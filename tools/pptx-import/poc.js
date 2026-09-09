const { JSDOM } = require("jsdom"); const { window } = new JSDOM(""); global.DOMParser = window.DOMParser;
// PoC：用 pptx-viewer 解析 sample.pptx，验证输出可映射到本项目 IR
const fs = require('fs');
const path = require('path');
const { loadPresentation } = require('pptx-viewer');

(async () => {
  const buf = fs.readFileSync(path.join(__dirname, '../../fixtures/sample.pptx'));
  const pres = await loadPresentation(buf);
  console.log('slideSize:', JSON.stringify(pres.slideSize));
  console.log('metadata:', JSON.stringify(pres.metadata));
  console.log('slides:', pres.slides.length);
  console.log('theme colors:', JSON.stringify(pres.theme && pres.theme.colors));
  console.log('theme fonts:', JSON.stringify(pres.theme && pres.theme.fonts));

  for (const s of pres.slides) {
    console.log('\n===== slide', s.index, 'bg:', JSON.stringify(s.background));
    for (const el of s.elements) {
      const b = el.bounds;
      console.log(`  [${el.type}] id=${el.id} bounds=${JSON.stringify(b)} rot=${el.rotation}`);
      if (el.type === 'text' || (el.type === 'shape' && el.text)) {
        const tb = el.text;
        console.log('    verticalAlign:', tb.verticalAlign, 'padding:', JSON.stringify(tb.padding));
        for (const p of tb.paragraphs) {
          const runDesc = (p.runs || []).map(r => {
            const parts = [JSON.stringify(r.text.slice(0, 18))];
            if (r.fontSize) parts.push('s=' + r.fontSize);
            if (r.fontFamily) parts.push('f=' + r.fontFamily);
            if (r.color) parts.push('c=' + r.color.hex + (r.color.alpha !== undefined && r.color.alpha !== 1 ? '@' + r.color.alpha : ''));
            if (r.bold) parts.push('b');
            if (r.italic) parts.push('i');
            if (r.underline) parts.push('u');
            return parts.join(' ');
          }).join(' | ');
          console.log('    p:', JSON.stringify({ al: p.alignment, lh: p.lineHeight, lsu: p.lineSpacingUnit, sb: p.spaceBefore, sa: p.spaceAfter, runs: runDesc }));
        }
        if (el.type === 'shape') console.log('    shapeType:', el.shapeType, 'fill:', JSON.stringify(el.fill), 'adjust:', el.adjustments && Array.from(el.adjustments.entries()));
      } else if (el.type === 'table') {
        console.log('    table cols:', JSON.stringify(el.columns && el.columns.map(c => c.width)));
        for (const row of el.rows) {
          console.log('    row h=' + (row.height || '?'), 'cells:', row.cells.map(c => JSON.stringify({ txt: c.text && c.text.paragraphs && c.text.paragraphs.map(p => p.runs.map(r => r.text).join('')).join('/'), fill: c.fill && c.fill.type, v: c.verticalAlign })).join(' ; '));
        }
      } else if (el.type === 'image') {
        console.log('    src len:', el.src.length, 'mime:', el.mimeType, 'head:', el.src.slice(0, 30));
      } else if (el.type === 'group') {
        console.log('    group children:', (el.elements || []).length);
      } else if (el.type === 'chart' || el.type === 'diagram') {
        console.log('    (chart/diagram, 降级为图片或提示)');
      }
    }
  }
  pres.cleanup && pres.cleanup();
})().catch(e => { console.error('FAIL', e); process.exit(1); });
