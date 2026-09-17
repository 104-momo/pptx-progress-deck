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
    /* 字体不在这里决定，由 pptx2ir 末尾的 fixFonts 统一处理（中文走东亚字体链） */
    if(r.eaFont) ro._eaFont = r.eaFont;
    if(r.fontFamily) ro._latinFont = r.fontFamily;
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

  /* 从第一性原理处理中文字体：
   * PPT 中文字体继承链：主题 → 母版 → 版式 → 幻灯片 → 文本框默认 → 段落默认 → run
   * pptx-viewer 只完整实现了 latin 字体继承，ea 字体继承链断裂。
   * 策略：中文 run 永远走东亚字体链，不允许 fallback 到 latin 字体。
   * 优先级：run.eaFont（解析主题引用后）→ 主题 minorEA/majorEA → 默认"宋体" */
  const themeFonts = (pres.theme && pres.theme.fonts) || null;
  const defaultEA = (themeFonts && (themeFonts.minorEA || themeFonts.majorEA)) || '宋体';

  const resolveThemeRef = function(f){
    if(!f) return f;
    if(f === '+mj-ea') return (themeFonts && themeFonts.majorEA) || defaultEA;
    if(f === '+mn-ea') return (themeFonts && themeFonts.minorEA) || defaultEA;
    if(f === '+mj-lt') return themeFonts ? themeFonts.major : f;
    if(f === '+mn-lt') return themeFonts ? themeFonts.minor : f;
    return f;
  };

  /* 常见中文字体名（用于判断 fontFamily 是否本身就是中文字体） */
  const CJK_FONTS = ['宋体','SimSun','Songti','微软雅黑','Microsoft YaHei','PingFang','黑体','SimHei','Heiti','楷体','KaiTi','Kaiti','仿宋','FangSong','STFangsong','等线','DengXian','幼圆','YouYuan','隶书','LiSu','华文','ST'];
  const isCJKFont = function(f){
    if(!f) return false;
    return CJK_FONTS.some(function(n){ return f.indexOf(n) >= 0; });
  };

  const fixFonts = function(paras){
    (paras || []).forEach(p => {
      (p.runs || []).forEach(r => {
        if(!r.tx) return;
        const hasCJK = /[\u4e00-\u9fff]/.test(r.tx);
        const eaFont = r._eaFont;
        const latinFont = r._latinFont;
        delete r._eaFont;
        delete r._latinFont;
        if(hasCJK){
          /* 中文：优先 eaFont（但必须是中文字体名，西文字体名如Calibri是PowerPoint误写，忽略），
             其次 latinFont（如果本身是中文字体），最后默认东亚字体 */
          if(eaFont && isCJKFont(eaFont)){
            r.f = resolveThemeRef(eaFont);
          }else if(latinFont && isCJKFont(latinFont)){
            r.f = resolveThemeRef(latinFont);
          }else{
            r.f = defaultEA;
          }
        }else if(latinFont){
          /* 非中文：用 latin 字体 */
          r.f = resolveThemeRef(latinFont);
        }
      });
    });
  };

  slides.forEach(s => {
    (s.shapes || []).forEach(sh => {
      fixFonts(sh.p);
      if(sh.rows){
        sh.rows.forEach(row => {
          (row.cells || []).forEach(cell => fixFonts(cell.p));
        });
      }
    });
  });

  /* 尺寸转换：pptx-viewer 输出的是 96DPI 像素，渲染器 slide 固定 1280px 宽，需统一缩放 */
  const scale = 1280 / pres.slideSize.width;
  const DEFAULT_FONT_PX = 24;  /* 18pt @96DPI，PowerPoint 正文默认字号 */
  const convertParas = function(paras){
    (paras || []).forEach(p => {
      if(p.sb) p.sb = Math.round(p.sb * scale);
      if(p.sa) p.sa = Math.round(p.sa * scale);
      if(p.lh && p.lsu === 'px') p.lh = Math.round(p.lh * scale);
      (p.runs || []).forEach(r => {
        if(r.s) r.s = Math.round(r.s * scale);
        else if(r.tx) r.s = Math.round(DEFAULT_FONT_PX * scale);
      });
    });
  };
  slides.forEach(s => {
    (s.shapes || []).forEach(sh => {
      if(sh.x != null) sh.x = Math.round(sh.x * scale);
      if(sh.y != null) sh.y = Math.round(sh.y * scale);
      if(sh.w != null) sh.w = Math.round(sh.w * scale);
      if(sh.h != null) sh.h = Math.round(sh.h * scale);
      if(sh.radius) sh.radius = Math.round(sh.radius * scale);
      convertParas(sh.p);
      if(sh.rows){
        if(sh.colW) sh.colW = sh.colW.map(function(w){ return Math.round(w * scale); });
        sh.rows.forEach(row => {
          if(row.h) row.h = Math.round(row.h * scale);
          (row.cells || []).forEach(cell => convertParas(cell.p));
        });
      }
    });
  });

  return {
    id: (opt && opt.id) || 'deck',
    title: title,
    size: { w: 1280, h: Math.round(pres.slideSize.height * scale) },
    slides: slides,
    skipped: skipped
  };
}
