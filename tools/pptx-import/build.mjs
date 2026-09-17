/* esbuild 打包：浏览器端 PPTX 导入器 → dist/pptx-import.js（被 render.py 内联进产物 HTML） */
import { build } from 'esbuild';
import fs from 'fs';

/* Patch pptx-viewer：原生只取 a:latin/@typeface，补上 a:ea/@typeface（东亚字体，如宋体）
   共四处：run级、主题解析、主题替换、段落继承 */
const patchPptxViewer = {
  name: 'patch-pptx-viewer-font',
  setup(build){
    build.onLoad({ filter: /pptx-viewer\.js$/ }, async (args) => {
      let src = fs.readFileSync(args.path, 'utf8');
      let patched = 0;

      /* Patch 1: run 级字体解析 —— 在 latin 提取后补 ea 提取 */
      const old1 = `  const u = h(e, "latin");
  if (u) {
    const g = m(u, "typeface");
    g && (r.fontFamily = g);
  }`;
      const rep1 = `  const u = h(e, "latin");
  if (u) {
    const g = m(u, "typeface");
    g && (r.fontFamily = g);
  }
  const eaEl = h(e, "ea");
  if (eaEl) {
    const ge = m(eaEl, "typeface");
    ge && (r.eaFont = ge);
  }`;
      if(src.includes(old1)){ src = src.replace(old1, rep1); patched++; }

      /* Patch 2: 主题字体解析 —— majorFont/minorFont 也提取 ea typeface */
      const old2 = `function Fn(e) {
  const t = {
    major: "Calibri Light",
    minor: "Calibri"
  }, n = v(e, "fontScheme");
  if (!n) return t;
  const r = v(n, "majorFont");
  if (r) {
    const i = v(r, "latin");
    if (i) {
      const l = m(i, "typeface");
      l && (t.major = l);
    }
  }
  const s = v(n, "minorFont");
  if (s) {
    const i = v(s, "latin");
    if (i) {
      const l = m(i, "typeface");
      l && (t.minor = l);
    }
  }
  return t;
}`;
      const rep2 = `function Fn(e) {
  const t = {
    major: "Calibri Light",
    minor: "Calibri",
    majorEA: "",
    minorEA: ""
  }, n = v(e, "fontScheme");
  if (!n) return t;
  const r = v(n, "majorFont");
  if (r) {
    const i = v(r, "latin");
    if (i) {
      const l = m(i, "typeface");
      l && (t.major = l);
    }
    const ie = v(r, "ea");
    if (ie) {
      const le = m(ie, "typeface");
      le && (t.majorEA = le);
    }
  }
  const s = v(n, "minorFont");
  if (s) {
    const i = v(s, "latin");
    if (i) {
      const l = m(i, "typeface");
      l && (t.minor = l);
    }
    const is = v(s, "ea");
    if (is) {
      const ls = m(is, "typeface");
      ls && (t.minorEA = ls);
    }
  }
  return t;
}`;
      if(src.includes(old2)){ src = src.replace(old2, rep2); patched++; }

      /* Patch 3: 主题字体替换 —— s() 函数也处理 +mj-ea/+mn-ea，遍历 runs 时也替换 eaFont */
      const old3 = `  const s = (l) => l.startsWith("+mj-") ? r.fonts.major : l.startsWith("+mn-") ? r.fonts.minor : l, i = (l) => {`;
      const rep3 = `  const s = (l) => {
    if (l.startsWith("+mj-ea")) return r.fonts.majorEA || r.fonts.major;
    if (l.startsWith("+mn-ea")) return r.fonts.minorEA || r.fonts.minor;
    if (l.startsWith("+mj-")) return r.fonts.major;
    if (l.startsWith("+mn-")) return r.fonts.minor;
    return l;
  }, i = (l) => {`;
      if(src.includes(old3)){ src = src.replace(old3, rep3); patched++; }

      /* Patch 3b: 遍历 runs 时，eaFont 也做主题替换 —— 不做，在 pptx2ir.js 中兜底 */

      /* Patch 4: 段落字体继承 —— run 从段落继承字体时也继承 eaFont（两处） */
      const old4a = `c.fontSize === void 0 && l.fontSize !== void 0 && (c.fontSize = l.fontSize), c.fontFamily === void 0 && l.fontFamily !== void 0 && (c.fontFamily = l.fontFamily),`;
      const rep4a = `c.fontSize === void 0 && l.fontSize !== void 0 && (c.fontSize = l.fontSize), c.fontFamily === void 0 && l.fontFamily !== void 0 && (c.fontFamily = l.fontFamily), c.eaFont === void 0 && l.eaFont !== void 0 && (c.eaFont = l.eaFont),`;
      if(src.includes(old4a)){ src = src.replace(old4a, rep4a); patched++; }

      const old4b = `i.fontSize === void 0 && s.fontSize !== void 0 && (i.fontSize = s.fontSize), i.fontFamily === void 0 && s.fontFamily !== void 0 && (i.fontFamily = s.fontFamily),`;
      const rep4b = `i.fontSize === void 0 && s.fontSize !== void 0 && (i.fontSize = s.fontSize), i.fontFamily === void 0 && s.fontFamily !== void 0 && (i.fontFamily = s.fontFamily), i.eaFont === void 0 && s.eaFont !== void 0 && (i.eaFont = s.eaFont),`;
      if(src.includes(old4b)){ src = src.replace(old4b, rep4b); patched++; }

      /* Patch 5: 文本框级别默认字号 —— ve() 解析 txBody/lstStyle 中的 lvlNpPr/defRPr */
      const old5 = `  return {
    paragraphs: r,
    verticalAlign: l,
    padding: i,
    autofit: c
  };
}`;
      const rep5 = `  const _ls = h(e, "lstStyle");
  if (_ls) {
    const _df = Nt(_ls, t);
    for (const _pa of r) {
      const _lv = _pa.level ?? 0, _d = _df[Math.min(_lv, _df.length - 1)];
      if (_d) for (const _rn of _pa.runs) {
        _rn.fontSize === void 0 && _d.fontSize !== void 0 && (_rn.fontSize = _d.fontSize);
        _rn.fontFamily === void 0 && _d.fontFamily !== void 0 && (_rn.fontFamily = _d.fontFamily);
        _rn.eaFont === void 0 && _d.eaFont !== void 0 && (_rn.eaFont = _d.eaFont);
        _rn.color === void 0 && _d.color !== void 0 && (_rn.color = _d.color);
        _rn.bold === void 0 && _d.bold !== void 0 && (_rn.bold = _d.bold);
        _rn.italic === void 0 && _d.italic !== void 0 && (_rn.italic = _d.italic);
        _rn.underline === void 0 && _d.underline !== void 0 && (_rn.underline = _d.underline);
      }
    }
  }
  return {
    paragraphs: r,
    verticalAlign: l,
    padding: i,
    autofit: c
  };
}`;
      if(src.includes(old5)){ src = src.replace(old5, rep5); patched++; }

      /* Patch 6: 段落级别默认字号 —— kn() 解析 pPr/defRPr 并继承到 run */
      const old6 = `  return {
    runs: r,
    align: i,
    lineSpacing: l,
    spaceBefore: c,
    spaceAfter: o,
    bullet: a,
    level: u,
    marginLeft: d,
    indent: f
  };
}`;
      const rep6 = `  if (s) {
    const _dr = h(s, "defRPr");
    if (_dr) {
      const _d = Ut(_dr, t);
      for (const _rn of r) {
        _rn.fontSize === void 0 && _d.fontSize !== void 0 && (_rn.fontSize = _d.fontSize);
        _rn.fontFamily === void 0 && _d.fontFamily !== void 0 && (_rn.fontFamily = _d.fontFamily);
        _rn.eaFont === void 0 && _d.eaFont !== void 0 && (_rn.eaFont = _d.eaFont);
        _rn.color === void 0 && _d.color !== void 0 && (_rn.color = _d.color);
        _rn.bold === void 0 && _d.bold !== void 0 && (_rn.bold = _d.bold);
        _rn.italic === void 0 && _d.italic !== void 0 && (_rn.italic = _d.italic);
        _rn.underline === void 0 && _d.underline !== void 0 && (_rn.underline = _d.underline);
      }
    }
  }
  return {
    runs: r,
    align: i,
    lineSpacing: l,
    spaceBefore: c,
    spaceAfter: o,
    bullet: a,
    level: u,
    marginLeft: d,
    indent: f
  };
}`;
      if(src.includes(old6)){ src = src.replace(old6, rep6); patched++; }

      /* Patch 7: 占位符类型映射 —— obj 类型 fallback 到 body 样式（other 通常为空） */
      const old7 = `function ur(e, t) {
  return e === "title" || e === "ctrTitle" ? t.title : e === "body" || e === "subTitle" ? t.body : t.other;
}`;
      const rep7 = `function ur(e, t) {
  const _r = e === "title" || e === "ctrTitle" ? t.title : e === "body" || e === "subTitle" ? t.body : t.other;
  if (_r && _r.length > 0) return _r;
  return t.body && t.body.length > 0 ? t.body : _r;
}`;
      if(src.includes(old7)){ src = src.replace(old7, rep7); patched++; }

      console.log('  [patch] applied ' + patched + ' patches');
      return { contents: src, loader: 'js' };
    });
  }
};

await build({
  entryPoints: ['src/entry.js'],
  bundle: true,
  format: 'iife',
  target: ['es2017'],
  minify: true,
  outfile: 'dist/pptx-import.js',
  logLevel: 'info',
  plugins: [patchPptxViewer]
});

const kb = (fs.statSync('dist/pptx-import.js').size / 1024).toFixed(1);
console.log(`\ndist/pptx-import.js: ${kb} KB`);
