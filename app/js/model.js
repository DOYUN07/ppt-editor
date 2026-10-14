'use strict';
// ============================================================================
// Data model
//   deck  = { title, footer, size:{w,h}, slides:[slide] }
//   slide = { id, template:{type, data} | null, ov:{role:{...}}, del:[role], extra:[element], bg, transition, notes, hidden }
// A template slide generates elements from its data (auto layout). Every generated element has a `role`
// (usually the data path it shows, e.g. "cards.1.title"). User edits to generated elements are stored as
// overrides (ov) or deletions (del); elements the user adds are kept in `extra`. Text typed on the canvas
// is written back into the data, so changing the layout keeps the content.
// ============================================================================
const SW = 960, SH = 540;
const uid = () => 'e' + Math.random().toString(36).slice(2, 9);
const clone = o => JSON.parse(JSON.stringify(o));
const getPath = (o, p) => p.split('.').reduce((a, k) => (a == null ? a : a[k]), o);
const setPath = (o, p, v) => { const ks = p.split('.'); let a = o; ks.slice(0, -1).forEach((k, i) => { if (a[k] == null) a[k] = /^\d+$/.test(ks[i + 1]) ? [] : {}; a = a[k]; }); a[ks[ks.length - 1]] = v; };
const N2 = i => String(i).padStart(2, '0');

// ---------------- colors ----------------
const hx = h => String(h).replace('#', '').toUpperCase();
const toRgb = h => { h = hx(h); return [0, 2, 4].map(i => parseInt(h.substr(i, 2), 16)); };
const toHex = rgb => rgb.map(v => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0')).join('').toUpperCase();
const mix = (a, b, t) => { const x = toRgb(a), y = toRgb(b); return toHex(x.map((v, i) => v + (y[i] - v) * t)); };
const lum = h => { const c = toRgb(h).map(v => { v /= 255; return v <= .03928 ? v / 12.92 : Math.pow((v + .055) / 1.055, 2.4); }); return .2126 * c[0] + .7152 * c[1] + .0722 * c[2]; };
const contrast = (a, b) => { const l1 = lum(a), l2 = lum(b); return (Math.max(l1, l2) + .05) / (Math.min(l1, l2) + .05); };
function darkenUntil(c, against, ratio) { let out = hx(c), n = 0; while (contrast(out, against) < ratio && n < 40) { out = mix(out, '000000', 0.06); n++; } return { color: out, changed: n > 0 }; }
function lightenUntil(c, against, ratio) { let out = hx(c), n = 0; while (contrast(out, against) < ratio && n < 40) { out = mix(out, 'FFFFFF', 0.06); n++; } return out; }
// Derive the 12-color theme from the picked colors and fix contrast automatically.
function derivePalette(b) {
  const notes = [];
  const m = darkenUntil(b.main, 'FFFFFF', 7);
  if (m.changed) notes.push('메인 색이 밝아서 흰 글씨가 잘 보이도록 조금 어둡게 조정했어요.');
  const dk2 = m.color;
  const a = darkenUntil(b.accent, 'FFFFFF', 3);
  if (a.changed) notes.push('강조 색 위의 흰 글씨가 잘 보이도록 강조 색을 조금 진하게 했어요.');
  const accent1 = a.color;
  const accent2 = b.subAuto ? mix(dk2, accent1, 0.5) : darkenUntil(b.sub, 'FFFFFF', 3).color;
  const dk1 = darkenUntil(mix(dk2, '0B1020', 0.7), 'FFFFFF', 13).color;
  const lt2 = mix(dk2, 'FFFFFF', 0.95);
  const accent3 = lightenUntil(mix(dk2, 'FFFFFF', 0.68), dk2, 4.5);
  if (contrast(accent1, dk2) < 2) notes.push('메인 색과 강조 색이 비슷해서 구분이 약해요. 강조 색을 다른 계열로 고르면 좋아요.');
  if (!notes.length) notes.push('ok:색 대비가 충분해요.');
  return { colors: { dk1, lt1: 'FFFFFF', dk2, lt2, accent1, accent2, accent3, accent4: '10BF7F', accent5: 'F97316', accent6: '6B7280', hlink: dk2, folHlink: accent2 }, notes };
}
const THEME_TOKENS = ['dk2', 'accent1', 'accent2', 'accent3', 'lt2', 'dk1', 'accent6', 'lt1', 'accent4', 'accent5'];
// Theme colour with brightness, like PowerPoint's "lighter 40%": 'accent1@0.4' (toward white), 'dk2@-0.3' (toward black).
// Changing the brand colours recolours these too.
// A colour can also be stored relative to a theme colour as a hue / saturation / lightness difference:
// 'dk2~12,-0.1,0.25' — exact for the original colours, and follows the brand when it changes.
const THEME_COLOR_RE = /^(dk1|lt1|dk2|lt2|accent[1-6])(?:@(-?\d*\.?\d+)|~(-?\d*\.?\d+),(-?\d*\.?\d+),(-?\d*\.?\d+))?$/;
const isThemeColor = c => THEME_COLOR_RE.test(String(c || ''));
const shadeHex = (h, b) => b > 0 ? mix(h, 'FFFFFF', b) : b < 0 ? mix(h, '000000', -b) : hx(h);
function hsl(h) {
  const [r, g, b] = toRgb(h).map(v => v / 255), mx = Math.max(r, g, b), mn = Math.min(r, g, b), l = (mx + mn) / 2, d = mx - mn;
  if (!d) return [0, 0, l];
  const s = d / (1 - Math.abs(2 * l - 1));
  const hh = mx === r ? ((g - b) / d) % 6 : mx === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [(hh * 60 + 360) % 360, Math.min(1, s), l];
}
function hslHex(h, s, l) {
  h = ((h % 360) + 360) % 360; s = Math.max(0, Math.min(1, s)); l = Math.max(0, Math.min(1, l));
  const c = (1 - Math.abs(2 * l - 1)) * s, x = c * (1 - Math.abs((h / 60) % 2 - 1)), m = l - c / 2;
  const [r, g, b] = h < 60 ? [c, x, 0] : h < 120 ? [x, c, 0] : h < 180 ? [0, c, x] : h < 240 ? [0, x, c] : h < 300 ? [x, 0, c] : [c, 0, x];
  return toHex([(r + m) * 255, (g + m) * 255, (b + m) * 255]);
}
function resolveColor(c, pv) {
  const m = String(c || '').match(THEME_COLOR_RE); if (!m) return hx(c);
  if (m[3] != null) { const [h, s, l] = hsl(pv[m[1]]); return hslHex(h + +m[3], s + +m[4], l + +m[5]); }
  return shadeHex(pv[m[1]], +(m[2] || 0));
}
const TOKEN_NAMES = { dk1: '진한 글자', lt1: '흰색', dk2: '메인', lt2: '연한 배경', accent1: '강조', accent2: '보조', accent3: '연한 보조', accent4: '초록', accent5: '주황', accent6: '회색' };

// ---------------- text helpers ----------------
function textObj(str, o = {}) {
  const lines = String(str ?? '').replace(/\r\n?/g, '\n').split('\n');
  const paras = lines.map(l => {
    if (o.list) { const lvl = (l.match(/^>+/) || [''])[0].length; return { runs: [{ t: l.replace(/^>+/, '') }], level: lvl + 1, bullet: o.list === true ? 'disc' : o.list, ...(lvl ? { runsSize: o.subSize } : {}) }; }
    return { runs: [{ t: l }] };
  });
  paras.forEach(p => { if (p.runsSize) { p.runs[0].size = p.runsSize; delete p.runsSize; } });
  return { paras, size: o.size || 16, color: o.color || 'dk1', font: o.font || 'body', bold: !!o.bold, align: o.align || 'left', valign: o.valign || 'top', lineSpacing: o.ls || 1.1,
    ...(o.spaceAfter != null ? { spaceAfter: o.spaceAfter } : {}), ...(o.bulletColor ? { bulletColor: o.bulletColor } : {}), ...(o.field ? { field: o.field } : {}), ...(o.pad ? { pad: o.pad } : {}) };
}
const runsText = p => (p.runs || []).map(r => r.t).join('');
const plainText = tx => (tx && tx.paras ? tx.paras.map(runsText).join('\n') : '');
const listLines = tx => (tx && tx.paras ? tx.paras.filter(p => runsText(p).trim() !== '').map(p => '>'.repeat(Math.max(0, (p.level || 1) - 1)) + runsText(p)) : []);
function isRich(tx) { return !!(tx && tx.paras && tx.paras.some(p => (p.runs || []).some(r => Object.keys(r).some(k => k !== 't')) || (p.align && p.align !== tx.align))); }

// ---------------- defaults ----------------
const DEFAULT_LOGO = null;
const DEFAULT_BRAND = { name: '기본', main: '#1F3A5F', accent: '#E8613C', sub: '#5B7FA6', subAuto: true, headFont: 'Pretendard', bodyFont: 'Pretendard', logo: null };
const KNOBS = {
  cover: ['blocks', 'circles', 'split', 'minimal'], coverTitle: ['middle', 'low'], mirror: [false, true], titleAlign: ['left', 'center'], kicker: [false, true],
  motif: ['rounded', 'square', 'circle'], contentBg: ['lt1', 'lt2'], agenda: ['grid', 'list'], cards: ['outline', 'filled', 'stack'], stats: ['plain', 'tiles'], process: ['horizontal', 'vertical']
};
const STRUCTURE = ['cover', 'coverTitle', 'mirror', 'titleAlign', 'kicker', 'motif', 'agenda', 'cards', 'process'];
const TYPES = {
  cover: '표지', agenda: '목차', section: '섹션 구분', bullets: '글머리표 + 이미지', photo: '사진', cards: '카드',
  stats: '숫자 지표', chart: '차트', process: '단계·로드맵', compare: '비교', table: '표', closing: '마무리', blank: '빈 슬라이드', titleOnly: '제목만'
};
function today() { const d = new Date(); return `${d.getFullYear()}. ${String(d.getMonth() + 1).padStart(2, '0')}. ${String(d.getDate()).padStart(2, '0')}`; }
const NEW_DATA = {
  cover: () => ({ section: '표지', title: '프레젠테이션 제목을\n입력하세요', subtitle: '부제목 또는 핵심 메시지', meta: '발표자 이름  ·  소속  ·  ' + today() }),
  agenda: () => ({ title: '목차', items: [{ title: '현황 분석', desc: '현재 상황을 진단합니다' }, { title: '핵심 과제', desc: '해결할 문제를 정의합니다' }, { title: '추진 전략', desc: '실행 방안을 제시합니다' }, { title: '기대 효과', desc: '성과와 다음 단계를 정리합니다' }] }),
  section: () => ({ number: '01', title: '섹션 제목', desc: '이 섹션에서 다룰 내용을 소개합니다' }),
  bullets: () => ({ title: '핵심 메시지를 제목으로 작성하세요', bullets: ['첫 번째 요점', '>세부 설명', '두 번째 요점', '세 번째 요점'], imageText: '이미지 · 도표 영역\n사진을 첨부하거나 여기로 끌어다 놓으세요' }),
  photo: () => ({ title: '사진과 함께 전하는\n핵심 메시지', text: '사진이 보여주는 내용을 두세 문장으로 설명합니다.', caption: '사진 설명 또는 출처' }),
  cards: () => ({ title: '세 가지 핵심 포인트', lead: '메시지를 한 줄로 요약합니다', cards: [{ title: '첫 번째 포인트', desc: '핵심 내용을 두세 줄로 설명합니다.' }, { title: '두 번째 포인트', desc: '핵심 내용을 두세 줄로 설명합니다.' }, { title: '세 번째 포인트', desc: '핵심 내용을 두세 줄로 설명합니다.' }] }),
  stats: () => ({ title: '숫자로 보는 핵심 성과', stats: [{ value: '85%', label: '고객 만족도', note: '전년 대비 12%p 상승' }, { value: '2.4배', label: '매출 성장', note: '3년 누적 기준' }, { value: '120억', label: '투자 유치', note: '시리즈 B' }], insight: '숫자가 의미하는 바를 한 문장으로 정리합니다.' }),
  chart: () => ({ title: '매출은 꾸준히 성장했습니다', caption: '연도별 매출 (단위: 억 원)', chartType: 'column', seriesName: '매출', categories: ['2022년', '2023년', '2024년', '2025년', '2026년'], values: [42, 55, 68, 81, 100], highlight: 4, insights: [{ title: '지속 성장', desc: '5년 연속 성장했습니다' }, { title: '성장 가속', desc: '처음으로 100억 원을 넘었습니다' }] }),
  process: () => ({ title: '추진 단계별 로드맵', current: 1, steps: [{ title: '진단', desc: '문제 정의', when: 'Q1' }, { title: '설계', desc: '실행 계획 수립', when: 'Q2' }, { title: '실행', desc: '파일럿 후 확대', when: 'Q3' }, { title: '정착', desc: '성과 측정', when: 'Q4' }] }),
  compare: () => ({ title: '현재와 개선 방향 비교', left: { label: 'AS-IS', title: '현재 상황', items: ['수작업 중심 업무', '분산된 데이터'] }, right: { label: 'TO-BE', title: '개선 방향', items: ['자동화로 시간 단축', '통합 데이터 플랫폼'] } }),
  table: () => ({ title: '세부 실행 계획', headers: ['과제', '담당', '일정', '상태'], rows: [['시장 조사', '기획팀', '1월', '완료'], ['서비스 설계', '서비스팀', '2월', '진행 중'], ['파일럿', '운영팀', '3월', '예정']], statusColors: { '완료': 'accent4', '진행 중': 'accent1', '예정': 'accent6' } }),
  closing: () => ({ section: '마무리', title: '감사합니다', subtitle: '질문과 의견을 환영합니다', meta: '이름  ·  email@company.com' }),
  titleOnly: () => ({ title: '제목을 입력하세요' }),
  blank: () => ({})
};
function newSlide(type) { return { id: uid(), template: type === 'free' ? null : { type, data: NEW_DATA[type]() }, ov: {}, del: [], extra: [], notes: '' }; }
function defaultDeck() {
  return { title: '프레젠테이션 제목', footer: '회사명  |  프레젠테이션 제목', size: { w: SW, h: SH },
    slides: ['cover', 'agenda', 'section', 'bullets', 'photo', 'cards', 'stats', 'chart', 'process', 'compare', 'table', 'closing'].map(newSlide) };
}
// Older saved projects stored slides as plain {type, ...data}.
function migrateDeck(deck) {
  if (!deck.size) deck.size = { w: SW, h: SH };
  deck.slides = (deck.slides || []).map(s => {
    if (s.id && ('template' in s)) { s.ov = s.ov || {}; s.del = s.del || []; s.extra = s.extra || []; return s; }
    const { type, notes, ...data } = s;
    return { id: uid(), template: { type, data }, ov: {}, del: [], extra: [], notes: notes || '' };
  });
  return deck;
}

// ---------------- layout recommendation ----------------
const pick = a => a[Math.floor(Math.random() * a.length)];
function diffCount(a, b) { return STRUCTURE.filter(k => String(a[k]) !== String(b[k])).length; }
function recommendLayout(cur, history) {
  const recent = (history || []).slice(-4);
  let best = null, bestScore = -1;
  for (let i = 0; i < 300; i++) {
    const L = {}; for (const [k, vals] of Object.entries(KNOBS)) L[k] = pick(vals);
    L.radius = L.motif === 'circle' ? 0.5 : L.motif === 'square' ? 0 : pick([0.08, 0.12, 0.2, 0.28]);
    L.titleSize = pick([28, 28, 30, 32]);
    const dCur = cur ? diffCount(L, cur) : 9;
    const dHist = recent.length ? Math.min(...recent.map(h => diffCount(L, h))) : 9;
    const score = Math.min(dCur, dHist + 1) + Math.random() * 0.5;
    if (dCur >= 5 && score > bestScore) { best = L; bestScore = score; }
  }
  best = best || Object.assign({}, cur);
  best.name = '양식-' + Math.random().toString(36).slice(2, 6).toUpperCase();
  return best;
}

// ============================================================================
// Template generators: data -> elements (all geometry for a 960x540 slide)
// ============================================================================
function motifOf(L) { return L.motif === 'circle' ? { shape: 'ellipse' } : L.motif === 'square' ? { shape: 'rect' } : { shape: 'roundRect', radius: L.radius ?? 0.12 }; }
function T(role, x, y, w, h, text, o = {}) { const e = { kind: 'text', role, x, y, w, h, text: textObj(text, o) }; if (o.isTitle) e.isTitle = true; if (o.locked) e.locked = true; return e; }
function B(role, shape, x, y, w, h, fill, o = {}) {
  const e = { kind: 'shape', role, shape, x, y, w, h, fill: fill || null };
  if (o.radius != null) e.radius = o.radius; if (o.line) e.line = o.line; if (o.alpha != null) e.fillOpacity = o.alpha; if (o.text) e.text = o.text; if (o.shadow) e.shadow = true;
  return e;
}
function M(ctx, role, x, y, w, h, fill, o = {}) { const m = motifOf(ctx.L); return B(role, m.shape, x, y, w, h, fill, Object.assign({ radius: m.radius }, o)); }
function badge(ctx, role, num, x, y, s, fill, ink) {
  const m = motifOf(ctx.L);
  return B(role, m.shape, x, y, s, s, fill, { radius: m.shape === 'roundRect' ? Math.min(0.3, (ctx.L.radius ?? 0.12) * 1.8) : undefined, text: textObj(num, { size: Math.round(s * 0.36), color: ink, bold: true, align: 'center', valign: 'middle', pad: [0, 0, 0, 0] }) });
}
const MX = (ctx, x, w) => ctx.mirror ? SW - x - w : x;
function logoDims(logo, h, maxW, wKey, hKey) { let hh = logo[hKey] || h, ww = logo[wKey] || 0; if (!ww) { const a = logo.aspect || 5.4; ww = hh * a; if (ww > maxW) { ww = maxW; hh = ww / a; } } return [ww, hh]; }
function logoEls(ctx, onDark, x, y) {
  const lg = ctx.brand.logo; if (!lg) return [];
  const [w, h] = logoDims(lg, 26, 200, 'coverWidth', 'coverHeight');
  if (onDark && lg.white) return [{ kind: 'image', role: 'frame.logo', x, y, w, h, image: { src: lg.white, url: lg.whiteUrl, fit: 'contain' } }];
  const out = [];
  if (onDark) out.push(B('frame.logoPlate', 'roundRect', x - 10, y - 8, w + 20, h + 16, 'lt1', { radius: 0.3 }));
  out.push({ kind: 'image', role: 'frame.logo', x, y, w, h, image: { src: lg.color, url: lg.colorUrl, fit: 'contain' } });
  return out;
}
function frameEls(ctx, d) {
  const TY = ctx.kicker ? 46 : 34, TH = ctx.kicker ? 46 : 52, al = ctx.center ? 'center' : 'left', out = [];
  if (ctx.kicker && ctx.kickerText) out.push(T('frame.kicker', 48, 26, 864, 16, ctx.kickerText, { size: 11, color: 'accent1', bold: true, align: al }));
  out.push(T('title', 48, TY, 864, TH, d.title || '', { size: ctx.L.titleSize || 28, color: 'dk2', bold: true, font: 'head', valign: 'middle', align: al, isTitle: true }));
  return out;
}
function footerEls(ctx) {
  const lg = ctx.brand.logo, out = []; let footLeft = 62;
  if (lg) { const [w, h] = logoDims(lg, 13, 110, 'footerWidth', 'footerHeight'); footLeft = 48 + w + 14; out.push({ kind: 'image', role: 'frame.footLogo', x: 48, y: 515 - h / 2, w, h, image: { src: lg.color, url: lg.colorUrl, fit: 'contain' } }); }
  else out.push(M(ctx, 'frame.mark', 48, 508.5, 7, 7, 'accent1'));
  out.push(T('frame.footer', footLeft, 505, 560 - footLeft, 20, ctx.footer || '', { size: 9, color: 'accent6', valign: 'middle' }));
  out.push(T('frame.page', 852, 505, 60, 20, '', { size: 9, color: 'accent6', align: 'right', valign: 'middle', field: 'page' }));
  return out;
}
function img(role, src, url, x, y, w, h) { return { kind: 'image', role, x, y, w, h, image: { src, url, fit: 'cover' } }; }
function coverDeco(ctx) {
  const L = ctx.L, X = (x, w) => MX(ctx, x, w);
  if (L.cover === 'circles') return [B('deco.1', 'ellipse', X(600, 470), 170, 470, 470, 'accent2', { alpha: 0.75 }), B('deco.2', 'ellipse', X(770, 150), 70, 150, 150, 'accent1'), B('deco.3', 'ellipse', X(660, 100), 400, 100, 100, null, { line: { color: 'accent3', width: 1.5 } })];
  if (L.cover === 'split') return [B('deco.1', 'rect', X(620, 340), 0, 340, 540, 'accent1'), M(ctx, 'deco.2', X(680, 220), 130, 220, 220, null, { line: { color: 'lt1', width: 2 } }), M(ctx, 'deco.3', X(820, 100), 300, 100, 100, 'dk2')];
  if (L.cover === 'minimal') return [B('deco.1', 'rect', X(620, 340), 0, 340, 540, 'lt2'), M(ctx, 'deco.2', X(680, 200), 150, 200, 200, 'accent1'), M(ctx, 'deco.3', X(800, 110), 290, 110, 110, 'accent2', { alpha: 0.85 })];
  return [M(ctx, 'deco.1', X(650, 210), 96, 210, 210, 'accent1'), M(ctx, 'deco.2', X(748, 170), 236, 170, 170, 'accent2', { alpha: 0.85 }), M(ctx, 'deco.3', X(618, 96), 330, 96, 96, null, { line: { color: 'accent3', width: 1.5 } })];
}
const GEN = {
  cover(ctx, d) {
    const L = ctx.L, light = ctx.light, tw = (L.cover === 'split' || L.cover === 'minimal') ? 520 : 540, x0 = MX(ctx, 58, tw);
    const Tg = L.coverTitle === 'low' ? { ty: 250, sy: 392, sh: 50, my: 470 } : { ty: 170, sy: 314, sh: 60, my: 456 };
    const muted = light ? 'accent6' : 'accent3';
    let els = d.image ? [img('image', d.image, d.imageUrl, MX(ctx, 600, 360), 0, 360, 540), M(ctx, 'deco.photo', ctx.mirror ? 268 : 624, 448, 68, 68, 'accent1')] : coverDeco(ctx);
    els = els.concat(ctx.brand.logo ? logoEls(ctx, !light, x0, 56) : [M(ctx, 'frame.mark', x0, Tg.ty - 30, 16, 16, 'accent1')]);
    els.push(T('title', x0, Tg.ty, tw, 130, d.title, { size: 40, color: light ? 'dk2' : 'lt1', bold: true, font: 'head', valign: 'bottom', ls: 1.05, isTitle: true }));
    els.push(T('subtitle', x0, Tg.sy, tw, Tg.sh, d.subtitle || '', { size: 18, color: muted }));
    if (d.meta) els.push(T('meta', x0, Tg.my, tw, 20, d.meta, { size: 12, color: muted }));
    return { bg: light ? 'lt1' : 'dk2', elements: els };
  },
  section(ctx, d) {
    const light = ctx.light, X = (x, w) => MX(ctx, x, w), al = ctx.mirror ? 'right' : 'left';
    const deco = ctx.L.cover === 'circles' ? [B('deco.1', 'ellipse', X(760, 300), 300, 300, 300, 'accent2', { alpha: 0.75 }), B('deco.2', 'ellipse', X(840, 80), 250, 80, 80, 'accent1')]
      : [M(ctx, 'deco.1', X(790, 120), 330, 120, 120, 'accent1'), M(ctx, 'deco.2', X(860, 70), 400, 70, 70, 'accent2')];
    return { bg: light ? 'lt2' : 'dk2', elements: deco.concat([
      T('number', X(58, 300), 150, 300, 96, d.number || '01', { size: 72, color: 'accent1', bold: true, valign: 'bottom', align: al }),
      T('title', X(58, 640), 250, 640, 70, d.title || '', { size: 36, color: light ? 'dk2' : 'lt1', bold: true, font: 'head', valign: 'bottom', align: al, isTitle: true }),
      T('desc', X(58, 640), 332, 640, 50, d.desc || '', { size: 16, color: light ? 'accent6' : 'accent3', align: al })]) };
  },
  agenda(ctx, d) {
    const items = d.items || [], els = frameEls(ctx, d);
    if (ctx.L.agenda === 'list') {
      const rh = Math.min(68, 340 / Math.max(1, items.length));
      items.forEach((it, i) => { const y = 124 + i * rh;
        els.push(T(`items.${i}.num`, 48, y, 80, rh, N2(i + 1), { size: 28, color: 'accent1', bold: true, valign: 'middle' }), T(`items.${i}.title`, 136, y, 320, rh, it.title, { size: 20, color: 'dk2', bold: true, valign: 'middle' }));
        if (it.desc) els.push(T(`items.${i}.desc`, 470, y, 442, rh, it.desc, { size: 14, color: 'accent6', valign: 'middle' }));
        els.push(B(`items.${i}.line`, 'rect', 48, y + rh - 0.5, 864, 1, '#DCDFE6', {})); });
    } else {
      const rows = Math.ceil(items.length / 2), bh = Math.min(126, (340 - (rows - 1) * 20) / Math.max(1, rows));
      items.forEach((it, i) => { const x = 48 + (i % 2) * 444, y = 130 + Math.floor(i / 2) * (bh + 20);
        els.push(B(`items.${i}.card`, 'roundRect', x, y, 420, bh, ctx.card, { radius: 0.06 }), badge(ctx, `items.${i}.num`, N2(i + 1), x + 28, y + (bh - 48) / 2, 48, 'dk2', 'lt1'));
        if (it.desc) els.push(T(`items.${i}.title`, x + 96, y + bh / 2 - 26, 300, 26, it.title, { size: 20, color: 'dk2', bold: true }), T(`items.${i}.desc`, x + 96, y + bh / 2 + 6, 296, 40, it.desc, { size: 14, color: 'accent6' }));
        else els.push(T(`items.${i}.title`, x + 96, y, 300, bh, it.title, { size: 20, color: 'dk2', bold: true, valign: 'middle' })); });
    }
    return { elements: els.concat(footerEls(ctx)) };
  },
  bullets(ctx, d) {
    const hasVis = d.image || d.imageText, bx = hasVis && ctx.mirror ? 492 : 48, vx = ctx.mirror ? 48 : 508, w = hasVis ? 420 : 864;
    const els = frameEls(ctx, d);
    els.push(T('bullets', bx, 124, w, 340, (d.bullets || []).join('\n'), { list: ctx.L.motif === 'circle' ? 'disc' : 'square', size: 16, subSize: 14, color: 'dk1', bulletColor: 'accent1', spaceAfter: 8, ls: 1.15 }));
    if (d.image) els.push(img('image', d.image, d.imageUrl, vx, 124, 404, 340));
    else if (d.imageText) {
      const tx = textObj(d.imageText, { size: 14, color: 'accent6', align: 'center', valign: 'middle' });
      if (tx.paras[0]) tx.paras[0].runs[0] = Object.assign(tx.paras[0].runs[0], { b: true, color: 'dk2' });
      els.push(B('imageText', 'roundRect', vx, 124, 404, 340, ctx.card, { radius: 0.04, line: { color: 'accent3', width: 1, dash: 'dash' }, text: tx }));
    }
    return { elements: els.concat(footerEls(ctx)) };
  },
  photo(ctx, d) {
    const ix = MX(ctx, 480, 480), tx0 = ctx.mirror ? 528 : 48;
    const els = [d.image ? img('image', d.image, d.imageUrl, ix, 0, 480, 540) : B('imageBox', 'rect', ix, 0, 480, 540, 'lt2', { text: textObj('사진을 첨부하거나 여기로 끌어다 놓으세요', { size: 14, color: 'accent6', align: 'center', valign: 'middle' }) })];
    els.push(M(ctx, 'deco.dot', tx0, 112, 14, 14, 'accent1'));
    els.push(T('title', tx0, 136, 384, 110, d.title || '', { size: ctx.L.titleSize || 28, color: 'dk2', bold: true, font: 'head', valign: 'bottom', isTitle: true }));
    if (d.text) els.push(T('text', tx0, 262, 384, 170, d.text, { size: 16, color: 'dk1', ls: 1.2 }));
    if (d.caption) els.push(T('caption', tx0, 470, 384, 20, d.caption, { size: 11, color: 'accent6' }));
    return { elements: els };
  },
  cards(ctx, d) {
    const cards = d.cards || [], n = Math.max(1, cards.length), TY = ctx.kicker ? 46 : 34, TH = ctx.kicker ? 46 : 52, els = frameEls(ctx, d);
    if (d.lead) els.push(T('lead', 48, TY + TH + 4, 864, 22, d.lead, { size: 14, color: 'accent6', align: ctx.center ? 'center' : 'left' }));
    if (ctx.L.cards === 'stack') {
      const rh = Math.min(100, (330 - (n - 1) * 14) / n);
      cards.forEach((c, i) => { const y = 136 + i * (rh + 14);
        els.push(B(`cards.${i}.card`, 'roundRect', 48, y, 864, rh, ctx.card, { radius: 0.1 }), badge(ctx, `cards.${i}.num`, N2(i + 1), 72, y + (rh - 44) / 2, 44, 'dk2', 'lt1'),
          T(`cards.${i}.title`, 140, y, 280, rh, c.title, { size: 18, color: 'dk2', bold: true, valign: 'middle' }), T(`cards.${i}.desc`, 440, y, 448, rh, c.desc, { size: 14, color: 'accent6', valign: 'middle' })); });
    } else {
      const w = (864 - (n - 1) * 24) / n, filled = ctx.L.cards === 'filled';
      cards.forEach((c, i) => { const x = 48 + i * (w + 24);
        els.push(B(`cards.${i}.card`, 'roundRect', x, 140, w, 280, filled ? 'dk2' : 'lt1', { radius: 0.05, line: filled || ctx.L.contentBg === 'lt2' ? null : { color: '#DDE0E6', width: 1 } }),
          badge(ctx, `cards.${i}.num`, N2(i + 1), x + 28, 168, 52, filled ? 'accent1' : 'dk2', 'lt1'),
          T(`cards.${i}.title`, x + 28, 246, w - 56, 28, c.title, { size: 20, color: filled ? 'lt1' : 'dk2', bold: true }),
          T(`cards.${i}.desc`, x + 28, 286, w - 56, 120, c.desc, { size: 14, color: filled ? 'accent3' : 'accent6' })); });
    }
    return { elements: els.concat(footerEls(ctx)) };
  },
  stats(ctx, d) {
    const st = d.stats || [], n = Math.max(1, st.length), w = (864 - (n - 1) * 24) / n, vs = n >= 4 ? 44 : 54, els = frameEls(ctx, d);
    st.forEach((s, i) => { const x = 48 + i * (w + 24);
      if (ctx.L.stats === 'tiles') {
        els.push(B(`stats.${i}.tile`, 'roundRect', x, 120, w, 180, ctx.card, { radius: 0.06 }), T(`stats.${i}.value`, x + 24, 136, w - 48, 70, s.value, { size: vs - 6, color: 'accent1', bold: true, valign: 'bottom' }),
          T(`stats.${i}.label`, x + 24, 216, w - 48, 26, s.label, { size: 18, color: 'dk2', bold: true }));
        if (s.note) els.push(T(`stats.${i}.note`, x + 24, 246, w - 48, 40, s.note, { size: 14, color: 'accent6' }));
      } else {
        els.push(T(`stats.${i}.value`, x, 130, w, 80, s.value, { size: vs, color: 'dk2', bold: true, valign: 'bottom' }), T(`stats.${i}.label`, x, 222, w, 26, s.label, { size: 18, color: 'dk1', bold: true }));
        if (s.note) els.push(T(`stats.${i}.note`, x, 252, w, 22, s.note, { size: 14, color: 'accent6' }));
        if (i) els.push(B(`stats.${i}.sep`, 'rect', x - 12.5, 140, 1, 130, '#DCDFE6', {}));
      } });
    if (d.insight) els.push(B('insight.box', 'roundRect', 48, 330, 864, 130, 'dk2', { radius: 0.06 }), T('insightLabel', 80, 356, 400, 20, d.insightLabel || 'KEY INSIGHT', { size: 12, color: 'accent3', bold: true }), T('insight', 80, 384, 800, 56, d.insight, { size: 18, color: 'lt1', bold: true }));
    return { elements: els.concat(footerEls(ctx)) };
  },
  chart(ctx, d) {
    const ins = d.insights || [], els = frameEls(ctx, d); let cw = 864, cx = 40, ix = 612;
    if (ins.length) { cw = 540; if (ctx.mirror) { ix = 48; cx = 372; } }
    if (d.caption) els.push(T('caption', cx + 8, 112, cw, 20, d.caption, { size: 13, color: 'dk1', bold: true }));
    ins.forEach((it, i) => { const y = 140 + i * (330 / Math.max(3, ins.length));
      els.push(M(ctx, `insights.${i}.dot`, ix, y + 5, 12, 12, 'accent1'), T(`insights.${i}.title`, ix + 24, y, 276, 24, it.title, { size: 17, color: 'dk2', bold: true }));
      if (it.desc) els.push(T(`insights.${i}.desc`, ix + 24, y + 30, 276, 50, it.desc, { size: 14, color: 'accent6' })); });
    const series = d.series && d.series.length ? d.series : [{ name: d.seriesName || '값', values: d.values || [] }];
    els.push({ kind: 'chart', role: 'chart', x: cx, y: 140, w: cw + 8, h: 330, chart: { type: d.chartType || 'column', categories: d.categories || [], series, highlight: d.highlight ?? -1, showValues: true } });
    return { elements: els.concat(footerEls(ctx)) };
  },
  process(ctx, d) {
    const st = d.steps || [], n = Math.max(1, st.length), cur = Number(d.current ?? -1), els = frameEls(ctx, d);
    if (ctx.L.process === 'vertical') {
      const rh = Math.min(84, 340 / n);
      if (st.length > 1) els.push(B('steps.line', 'rect', 73, 124 + rh / 2, 2, (st.length - 1) * rh, '#DCDFE6', {}));
      st.forEach((s, i) => { const y = 124 + i * rh;
        els.push(badge(ctx, `steps.${i}.num`, N2(i + 1), 50, y + (rh - 48) / 2, 48, i === cur ? 'accent1' : 'dk2', 'lt1'),
          T(`steps.${i}.step`, 124, y + rh / 2 - 24, 200, 16, 'STEP ' + N2(i + 1), { size: 11, color: 'accent1', bold: true }), T(`steps.${i}.title`, 124, y + rh / 2 - 6, 300, 28, s.title, { size: 20, color: 'dk2', bold: true }));
        if (s.desc) els.push(T(`steps.${i}.desc`, 440, y, 330, rh, s.desc, { size: 14, color: 'accent6', valign: 'middle' }));
        if (s.when) els.push(T(`steps.${i}.when`, 780, y, 132, rh, s.when, { size: 12, color: 'accent6', bold: true, valign: 'middle', align: 'right' }));
        if (i < st.length - 1) els.push(B(`steps.${i}.sep`, 'rect', 124, y + rh - 0.5, 788, 1, '#DCDFE6', {})); });
    } else {
      const colW = 864 / n; els.push(B('steps.line', 'rect', 78, 185, (n - 1) * colW, 2, '#DCDFE6', {})); let anyWhen = false;
      st.forEach((s, i) => { const x = 48 + i * colW;
        els.push(badge(ctx, `steps.${i}.num`, N2(i + 1), x, 156, 60, i === cur ? 'accent1' : 'dk2', 'lt1'), T(`steps.${i}.step`, x, 240, colW - 20, 18, 'STEP ' + N2(i + 1), { size: 11, color: 'accent1', bold: true }),
          T(`steps.${i}.title`, x, 262, colW - 20, 30, s.title, { size: 22, color: 'dk2', bold: true }));
        if (s.desc) els.push(T(`steps.${i}.desc`, x, 302, colW - 20, 60, s.desc, { size: 14, color: 'accent6' }));
        if (s.when) { anyWhen = true; els.push(T(`steps.${i}.when`, x, 386, colW - 20, 20, s.when, { size: 12, color: 'accent6', bold: true })); } });
      if (anyWhen) els.push(B('steps.sep', 'rect', 48, 370, 864, 1, '#DCDFE6', {}));
    }
    return { elements: els.concat(footerEls(ctx)) };
  },
  compare(ctx, d) {
    const L_ = d.left || {}, R_ = d.right || {}, els = frameEls(ctx, d);
    els.push(B('left.box', 'roundRect', 48, 120, 410, 350, ctx.card, { radius: 0.05 }), B('right.box', 'roundRect', 502, 120, 410, 350, 'dk2', { radius: 0.05 }),
      T('left.label', 80, 148, 340, 18, L_.label || 'AS-IS', { size: 12, color: 'accent6', bold: true }), T('left.title', 80, 170, 340, 32, L_.title || '', { size: 22, color: 'dk2', bold: true }),
      T('right.label', 534, 148, 340, 18, R_.label || 'TO-BE', { size: 12, color: 'accent3', bold: true }), T('right.title', 534, 170, 340, 32, R_.title || '', { size: 22, color: 'lt1', bold: true }),
      T('left.items', 80, 232, 350, 220, (L_.items || []).join('\n'), { list: 'disc', size: 16, color: 'dk1', bulletColor: 'accent6', spaceAfter: 18 }),
      T('right.items', 534, 232, 350, 220, (R_.items || []).join('\n'), { list: ctx.L.motif === 'circle' ? 'disc' : 'square', size: 16, color: 'lt1', bulletColor: 'accent1', spaceAfter: 18 }),
      B('arrow', 'ellipse', 462, 277, 36, 36, 'accent1', { text: textObj('→', { size: 16, bold: true, color: 'lt1', align: 'center', valign: 'middle', pad: [0, 0, 0, 0] }) }));
    return { elements: els.concat(footerEls(ctx)) };
  },
  table(ctx, d) {
    const hd = d.headers || [], rows = d.rows || [], nr = rows.length + 1, rh = Math.min(56, 340 / nr), nc = Math.max(1, hd.length);
    const first = nc <= 2 ? 432 : Math.round(864 * 0.4), rest = (864 - first) / Math.max(1, nc - 1);
    const els = frameEls(ctx, d);
    els.push({ kind: 'table', role: 'table', x: 48, y: 124, w: 864, h: rh * nr, table: { rows: [hd].concat(rows.map(r => hd.map((_, c) => r[c] ?? ''))), header: true, colW: hd.map((_, c) => c ? rest : first), fontSize: 14, headerFill: 'dk2', headerColor: 'lt1', bandFill: 'lt2', bodyFill: 'lt1', color: 'dk1', banded: true, firstColBold: true, statusColors: d.statusColors || {} } });
    return { elements: els.concat(footerEls(ctx)) };
  },
  titleOnly(ctx, d) { return { elements: frameEls(ctx, d).concat(footerEls(ctx)) }; },
  blank() { return { elements: [] }; }
};
GEN.closing = GEN.cover;

// Text roles that write back into template data (other roles keep their edits as overrides).
const LIST_ROLES = ['bullets', 'left.items', 'right.items'];
function isDataRole(role) { return LIST_ROLES.includes(role) || /^(title|subtitle|meta|number|desc|lead|insight|insightLabel|caption|text|imageText|left\.label|left\.title|right\.label|right\.title)$/.test(role) || /^(items|cards|stats|steps|insights)\.\d+\.(title|desc|value|label|note|when)$/.test(role); }

// ---------------- resolve: template + overrides + extras -> final elements ----------------
const GEOM = ['x', 'y', 'w', 'h', 'rot'];
function applyOverride(el, ov) {
  if (!ov) return el;
  const e = Object.assign({}, el);
  for (const k of ['x', 'y', 'w', 'h', 'rot', 'flipH', 'flipV', 'fill', 'fillOpacity', 'opacity', 'line', 'shadow', 'shape', 'radius', 'anim', 'z', 'locked', 'name']) if (k in ov) e[k] = clone(ov[k]);
  if (ov.image && e.image) e.image = Object.assign({}, e.image, ov.image);
  if (ov.chart && e.chart) e.chart = Object.assign({}, e.chart, ov.chart);
  if (ov.table && e.table) e.table = Object.assign({}, e.table, ov.table);
  const style = Object.assign({}, ov.textStyle || {}); delete style.paras;
  if (ov.text && ov.textForce) {
    // text typed into a non-data element (badge number, decoration) is always kept
    e.text = Object.assign({}, e.text || textObj('', { size: 18, align: 'center', valign: 'middle' }), style, { paras: clone(ov.text.paras) });
  } else if (e.text && (ov.textStyle || ov.text)) {
    e.text = Object.assign({}, e.text, style);
    // rich text override is used only while its plain text still matches the generated text
    if (ov.text && plainText(ov.text) === plainText(el.text)) e.text.paras = clone(ov.text.paras);
  }
  return e;
}
function slideCtx(state, idx) {
  const L = state.layout, brand = state.brand;
  let kickerText = '';
  for (let j = idx; j >= 0; j--) { const s = state.deck.slides[j]; if (s && s.template && s.template.type === 'section') { const d = s.template.data; kickerText = `${d.number || '01'}   ${String(d.title || '').split('\n')[0]}`; break; } }
  return { L, brand, footer: state.deck.footer, light: L.cover === 'minimal', mirror: String(L.mirror) === 'true', kicker: String(L.kicker) === 'true', center: L.titleAlign === 'center',
    card: L.contentBg === 'lt2' ? 'lt1' : 'lt2', kickerText };
}
function resolveSlide(state, idx) {
  const s = state.deck.slides[idx];
  let bg = 'lt1', els = [];
  if (s.template && GEN[s.template.type]) {
    const ctx = slideCtx(state, idx);
    const g = GEN[s.template.type](ctx, s.template.data || {});
    bg = g.bg || (state.layout.contentBg === 'lt2' ? 'lt2' : 'lt1');
    els = g.elements.filter(e => !(s.del || []).includes(e.role)).map((e, i) => Object.assign(applyOverride(e, (s.ov || {})[e.role]), { id: 'r:' + e.role }))
      .map((e, i) => (e.z == null ? Object.assign(e, { z: i }) : e));
  }
  const extras = (s.extra || []).map((e, i) => (e.z == null ? Object.assign({}, e, { z: 1000 + i }) : e));
  const all = els.concat(extras).sort((a, b) => a.z - b.z);
  return { bg: s.bg || { color: bg }, elements: all };
}

// ---------------- template switching (keep content when changing a slide's layout type) ----------------
function toGeneric(type, d) {
  const g = { title: d.title || '', text: d.text || d.desc || d.subtitle || d.lead || '', items: [], image: d.image, imageUrl: d.imageUrl };
  if (d.items) g.items = d.items.map(x => ({ title: x.title, desc: x.desc || '' }));
  if (d.cards) g.items = d.cards.map(x => ({ title: x.title, desc: x.desc || '' }));
  if (d.steps) g.items = d.steps.map(x => ({ title: x.title, desc: x.desc || '', when: x.when }));
  if (d.stats) g.items = d.stats.map(x => ({ title: x.label, desc: x.note || '', value: x.value }));
  if (d.insights && !g.items.length) g.items = d.insights.map(x => ({ title: x.title, desc: x.desc || '' }));
  if (d.bullets) g.items = d.bullets.filter(b => !b.startsWith('>')).map(b => ({ title: b, desc: '' }));
  if (d.left) g.items = [{ title: d.left.title, desc: (d.left.items || []).join(', ') }, { title: d.right.title, desc: (d.right.items || []).join(', ') }];
  if (d.headers && d.rows) g.items = d.rows.map(r => ({ title: r[0], desc: r.slice(1).join(' · ') }));
  return g;
}
function fromGeneric(type, g) {
  const d = NEW_DATA[type](); if (g.title) d.title = g.title;
  const it = g.items.length ? g.items : null;
  switch (type) {
    case 'cover': case 'closing': if (g.text) d.subtitle = g.text; if (g.image) { d.image = g.image; d.imageUrl = g.imageUrl; } break;
    case 'section': d.desc = g.text || ''; break;
    case 'agenda': if (it) d.items = it.slice(0, 6).map(x => ({ title: x.title, desc: x.desc })); break;
    case 'bullets': if (it) d.bullets = it.flatMap(x => [x.title].concat(x.desc ? ['>' + x.desc] : [])); else if (g.text) d.bullets = [g.text]; if (g.image) { d.image = g.image; d.imageUrl = g.imageUrl; } break;
    case 'photo': d.text = g.text || (it ? it.map(x => x.title).join(', ') : d.text); if (g.image) { d.image = g.image; d.imageUrl = g.imageUrl; } break;
    case 'cards': if (it) d.cards = it.slice(0, 4).map(x => ({ title: x.title, desc: x.desc })); if (g.text) d.lead = g.text; break;
    case 'stats': if (it) d.stats = it.slice(0, 4).map((x, i) => ({ value: x.value || d.stats[i % d.stats.length].value, label: x.title, note: x.desc })); break;
    case 'process': if (it) d.steps = it.slice(0, 5).map(x => ({ title: x.title, desc: x.desc, when: x.when || '' })); break;
    case 'chart': if (it) d.insights = it.slice(0, 3).map(x => ({ title: x.title, desc: x.desc })); break;
    case 'compare': if (it && it.length >= 2) { d.left.title = it[0].title; d.left.items = String(it[0].desc || '').split(/,\s*/).filter(Boolean); d.right.title = it[1].title; d.right.items = String(it[1].desc || '').split(/,\s*/).filter(Boolean); } break;
    case 'table': if (it) { d.headers = ['항목', '내용']; d.rows = it.map(x => [x.title, x.desc]); } break;
    case 'titleOnly': break;
  }
  return d;
}

// ---------------- engine document ----------------
function engineDoc(state, pal) {
  const hex = c => (c && pal[c] ? '#' + pal[c] : c);
  const slides = state.deck.slides.map((s, i) => {
    const r = resolveSlide(state, i);
    const elements = r.elements.map(e => {
      const o = clone(e); delete o.id; delete o.role; delete o.z; delete o.locked;
      if (o.image) delete o.image.url;
      // shaded theme colours ('accent1@0.4') go to the engine as plain RGB; bare tokens keep the theme link
      (function shade(x) { if (!x || typeof x !== 'object') return; for (const k of Object.keys(x)) { const v = x[k]; if (typeof v === 'string' && /[@~]/.test(v) && isThemeColor(v)) x[k] = '#' + resolveColor(v, pal); else if (typeof v === 'object') shade(v); } })(o);
      if (o.kind === 'icon') o.icon = { svg: iconSvg(o.icon.name, hex(o.icon.color || 'dk2'), o.icon.stroke) };
      return o;
    }).filter(o => !(o.kind === 'image' && !o.image.src));
    const out = { elements };
    if (r.bg.image) out.bg = { image: r.bg.image }; else if (r.bg.color && r.bg.color !== 'lt1') out.bg = { color: r.bg.color };
    if (s.transition && s.transition.effect && s.transition.effect !== 'none') out.transition = s.transition;
    if (s.notes) out.notes = s.notes;
    if (s.hidden) out.hidden = true;
    const t = s.template && s.template.type;
    if (t === 'cover') out.section = (s.template.data.section || '표지');
    else if (t === 'section') out.section = String(s.template.data.title || '섹션').split('\n')[0];
    else if (t === 'closing') out.section = (s.template.data.section || '마무리');
    return out;
  });
  return { size: state.deck.size || { w: SW, h: SH }, theme: { name: String(state.layout.name || 'custom').replace(/\*$/, ''), colors: pal, headFont: state.brand.headFont, bodyFont: state.brand.bodyFont }, slides };
}
