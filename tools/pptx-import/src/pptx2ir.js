/* pptx-viewer 的 Presentation → 本项目中间层 IR（deck JSON）
 * 与 tools/pptx2html.py 的降级策略保持一致：
 * 渐变/图片填充 → 首个色标纯色或透明；chart/diagram → 跳过并计数提示。 */
function hex(c){ return c && c.hex ? c.hex : null; }

function fillToColor(f){
  if(!f) return null;
  if(f.type === 'solid') return hex(f.color);
  if(f.type === 'gradient' || f.type === 'pattern'){
    if(f.stops && f.stops.length && f.stops[0].color) return hex(f.stops[0].color);
    return null;
  }
  return null;
}

function anchorOf(va){ return va === 'middle' ? 'middle' : (va === 'bottom' ? 'bottom' : 'top'); }

function para2ir(p){
  const out = {};
  if(p.align) out.al = p.align;
  if(p.lineSpacing){ out.lsu = 'x'; out.lh = p.lineSpacing; }
  if(p.spaceBefore) out.sb = Math.round(p.spaceBefore);
  if(p.spaceAfter) out.sa = Math.round(p.spaceAfter);
  out.runs = (p.runs || []).map(r => {
    const ro = {};
    if(r.text) ro.tx = r.text;
    if(r.fontSize) ro.s = Math.round(r.fontSize);
    if(r.fontFamily) ro.f = r.fontFamily;
    const c = r.color && hex(r.color);
    if(c) ro.c = c;
    if(r.bold) ro.b = true;
    if(r.italic) ro.i = true;
    if(r.underline) ro.u = true;
    return ro;
  });
  return out;
}

function textBody2ir(tb){
  const out = { p: (tb.paragraphs || []).map(para2ir) };
  const a = anchorOf(tb.verticalAlign);
  if(a !== 'top') out.anchor = a;
  return out;
}

function roundRectRadius(el, w, h){
  const adj = el.adjustments && el.adjustments.get ? el.adjustments.get('adj1') : null;
  if(typeof adj === 'number' && adj > 0 && adj < 1) return Math.round(adj * Math.min(w, h));
  return 0;
}

function slide2ir(s, idx){
  const shapes = [];
  const skipped = [];
  (function walk(list){
    (list || []).forEach(el => {
      const b = el.bounds || {};
      const x = Math.round(b.x || 0), y = Math.round(b.y || 0),
            w = Math.round(b.width || 0), h = Math.round(b.height || 0);
      switch(el.type){
        case 'text': {
          const o = { t:'t', x: x, y: y, w: w, h: h };
          const f = fillToColor(el.fill); if(f) o.fill = f;
          const tb = textBody2ir(el.text);
          if(tb.anchor) o.anchor = tb.anchor;
          o.p = tb.p;
          shapes.push(o);
          break;
        }
        case 'shape': {
          const hasText = el.text && el.text.paragraphs && el.text.paragraphs.length;
          if(hasText || el.fill){
            const o = { t:'t', x: x, y: y, w: w, h: h };
            const f = fillToColor(el.fill); if(f) o.fill = f;
            if(el.shapeType === 'roundRect'){
              const r = roundRectRadius(el, w, h);
              if(r) o.radius = r;
            }
            if(hasText){
              const tb = textBody2ir(el.text);
              if(tb.anchor) o.anchor = tb.anchor;
              o.p = tb.p;
            }else{
              o.p = [];
            }
            shapes.push(o);
          }else{
            skipped.push('shape');
          }
          break;
        }
        case 'table': {
          const o = { t:'tb', x: x, y: y, w: w };
          if(el.columnWidths && el.columnWidths.length) o.colW = el.columnWidths.map(Math.round);
          o.rows = (el.rows || []).map(r => {
            const ro = { cells: (r.cells || []).map(c => {
              const co = {};
              const f = fillToColor(c.fill); if(f) co.fill = f;
              co.p = c.text ? (c.text.paragraphs || []).map(para2ir) : [];
              if(c.colSpan > 1) co.cs = c.colSpan;
              return co;
            })};
            if(r.height) ro.h = Math.round(r.height);
            return ro;
          });
          shapes.push(o);
          break;
        }
        case 'image': {
          shapes.push({ t:'i', x: x, y: y, w: w, h: h, src: el.src || '', mime: el.mimeType || '' });
          break;
        }
        case 'group': walk(el.elements); break;
        default: skipped.push(el.type);
      }
    });
  })(s.elements);
  const bg = (s.background && fillToColor(s.background.fill)) || '#FFFFFF';
  return { n: idx + 1, bg: bg, shapes: shapes, skipped: skipped };
}

export function pptx2ir(pres, opt){
  const skipped = {};
  const slides = (pres.slides || []).map((s, i) => {
    const r = slide2ir(s, i);
    r.skipped.forEach(t => { skipped[t] = (skipped[t] || 0) + 1; });
    delete r.skipped;
    return r;
  });
  const title = (opt && opt.title) || pres.metadata.title || '未命名课件';
  return {
    id: (opt && opt.id) || 'deck',
    title: title,
    size: { w: Math.round(pres.slideSize.width), h: Math.round(pres.slideSize.height) },
    slides: slides,
    skipped: skipped
  };
}
