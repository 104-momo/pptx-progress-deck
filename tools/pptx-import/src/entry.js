/* 浏览器端全局入口：window.PPTXImport.fromFile(file, opt) → deck IR */
import { loadPresentation } from 'pptx-viewer';
import { pptx2ir } from './pptx2ir.js';

/* blob URL → dataURL（用 canvas 重绘；在 pres.cleanup() 之前调用，blob 仍有效） */
function blobImgToDataURL(src, mime){
  return new Promise(function(res){
    var img=new Image();
    img.onload=function(){
      try{
        var c=document.createElement('canvas');
        c.width=img.naturalWidth||1; c.height=img.naturalHeight||1;
        c.getContext('2d').drawImage(img,0,0);
        res(c.toDataURL(mime||'image/png'));
      }catch(e){ res(src); }
    };
    img.onerror=function(){ res(src); };
    img.src=src;
  });
}

window.PPTXImport = {
  async fromFile(fileOrBuf, opt){
    const pres = await loadPresentation(fileOrBuf);
    try{
      const deck = pptx2ir(pres, opt);
      /* 把 pptx-viewer 产生的 blob: 图片转 dataURL，确保保存后持久化、file:// 下可显示 */
      const imgs = [];
      (deck.slides||[]).forEach(function(s){
        (s.shapes||[]).forEach(function(sh){
          if(sh.t==='i'&&sh.src&&sh.src.indexOf('blob:')===0) imgs.push(sh);
        });
      });
      await Promise.all(imgs.map(function(sh){
        return blobImgToDataURL(sh.src, sh.mime).then(function(d){ sh.src=d; });
      }));
      return deck;
    }finally{
      try{ pres.cleanup(); }catch(e){}
    }
  }
};
