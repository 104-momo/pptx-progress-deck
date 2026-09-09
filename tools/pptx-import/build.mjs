/* esbuild 打包：浏览器端 PPTX 导入器 → dist/pptx-import.js（被 render.py 内联进产物 HTML） */
import { build } from 'esbuild';
import fs from 'fs';

await build({
  entryPoints: ['src/entry.js'],
  bundle: true,
  format: 'iife',
  target: ['es2017'],
  minify: true,
  outfile: 'dist/pptx-import.js',
  logLevel: 'info'
});

const kb = (fs.statSync('dist/pptx-import.js').size / 1024).toFixed(1);
console.log(`\ndist/pptx-import.js: ${kb} KB`);
