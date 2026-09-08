# -*- coding: utf-8 -*-
"""
build/<ID>.json + tools/deck_template.html  ->  <ID>-<标题>.html （单文件，离线可用）

用法:
    python3 tools/render.py --id L01 --name "导论"
"""
import argparse
import html
import json
import os
import re

here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SERIF = 'var(--serif)'
SANS = 'var(--sans)'


def esc(s):
    return html.escape(str(s), quote=True)


def fam(name):
    """pptx 字体名 -> CSS 字体链（补中文字体，保证中文不会掉到默认宋体）"""
    if not name:
        return SANS
    n = name.lower()
    if n in ('georgia', 'times new roman', '宋体', 'simsun', 'songti sc'):
        return "'%s','Songti SC','SimSun',serif" % name
    if n in ('微软雅黑', 'microsoft yahei'):
        return "'%s','PingFang SC',sans-serif" % name
    return "'%s','PingFang SC','Microsoft YaHei',sans-serif" % name


def run_html(r):
    st = []
    if r.get('s'):
        st.append('font-size:%.1fpx' % r['s'])
    f = r.get('f')
    if f:
        st.append('font-family:%s' % fam(f))
    if r.get('b'):
        st.append('font-weight:700')
    if r.get('i'):
        st.append('font-style:italic')
    if r.get('u'):
        st.append('text-decoration:underline')
    c = r.get('c')
    if c:
        st.append('color:%s' % c)
    txt = esc(r.get('tx') or '')
    if not txt:
        return ''
    return '<span style="%s">%s</span>' % (';'.join(st), txt)


def para_html(p):
    """段落。文本里的换行交给 white-space:pre-wrap 处理"""
    st = []
    al = p.get('al') or 'left'
    st.append('text-align:%s' % al)
    lh, lsu = p.get('lh'), p.get('lsu')
    base = max([r['s'] for r in p['runs'] if r.get('s')] or [0])
    if lh:
        st.append('line-height:%s' % (('%s' % lh) if lsu == 'x' else ('%.1fpx' % lh)))
    if p.get('sb'):
        st.append('margin-top:%.1fpx' % p['sb'])
    if p.get('sa'):
        st.append('margin-bottom:%.1fpx' % p['sa'])
    inner = ''.join(run_html(r) for r in p['runs'])
    if not inner.strip():
        # 空段落也要占位（撑出 PPT 里的行距）
        inner = '<span>&nbsp;</span>'
    return '<div class="p" style="%s">%s</div>' % (';'.join(st), inner)


def text_shape(sh):
    st = ['left:%.1fpx' % sh['x'], 'top:%.1fpx' % sh['y'],
          'width:%.1fpx' % sh['w'], 'height:%.1fpx' % sh['h']]
    if sh.get('fill'):
        st.append('background:%s' % sh['fill'])
    if sh.get('radius'):
        st.append('border-radius:%dpx' % sh['radius'])
    anchor = (sh.get('anchor') or 'TOP').lower()
    acls = {'middle': 'mid', 'top': 'top', 'bottom': 'bot'}.get(anchor, 'top')
    paras = ''.join(para_html(p) for p in sh.get('p', []))
    return '<div class="sh" style="%s"><div class="tf %s">%s</div></div>' % (
        ';'.join(st), acls, paras)


def table_shape(sh):
    st = ['left:%.1fpx' % sh['x'], 'top:%.1fpx' % sh['y'],
          'width:%.1fpx' % sh['w']]
    parts = ['<div class="sh" style="%s"><table class="dt">' % ';'.join(st)]
    if sh.get('colW'):
        parts.append('<colgroup>')
        for w in sh['colW']:
            parts.append('<col style="width:%.1fpx">' % w)
        parts.append('</colgroup>')
    for row in sh['rows']:
        parts.append('<tr>')
        for cell in row['cells']:
            cs = []
            if cell.get('fill'):
                cs.append('background:%s' % cell['fill'])
            if row.get('h'):
                cs.append('height:%.1fpx' % row['h'])
            inner = ''.join(para_html(p) for p in cell.get('p', []))
            parts.append('<td style="%s">%s</td>' % (';'.join(cs), inner))
        parts.append('</tr>')
    parts.append('</table></div>')
    return ''.join(parts)


def img_shape(sh):
    st = ['left:%.1fpx' % sh['x'], 'top:%.1fpx' % sh['y'],
          'width:%.1fpx' % sh['w'], 'height:%.1fpx' % sh['h']]
    return '<div class="sh" style="%s"><img class="shapeimg" src="%s" alt=""></div>' % (
        ';'.join(st), sh.get('src') or '')


def render(deck, tpl, classes=None, lectures=None, course="", course_id=""):
    slides = []
    for s in deck['slides']:
        body = []
        for sh in s['shapes']:
            t = sh['t']
            if t == 't':
                body.append(text_shape(sh))
            elif t == 'tb':
                body.append(table_shape(sh))
            elif t == 'i':
                body.append(img_shape(sh))
        slides.append('<section class="slide" data-n="%d" style="background:%s">%s</section>'
                      % (s['n'], s.get('bg') or '#FFFFFF', ''.join(body)))

    out = tpl
    total = len(deck['slides'])
    title = deck.get('title') or deck['id']
    for k, v in [('__LECTURE_TITLE_HTML__', esc(title)),
                 ('__LECTURE_TITLE__', json.dumps(title, ensure_ascii=False)),
                 ('__LECTURE_ID__', json.dumps(deck['id'], ensure_ascii=False)),
                 ('__LECTURE_ID_HTML__', esc(deck['id'])),
                 ('__COURSE__', json.dumps(course, ensure_ascii=False)),
                 ('__COURSE_ID__', json.dumps(course_id or 'deck', ensure_ascii=False)),
                 ('__CLASSES__', json.dumps(classes or [], ensure_ascii=False)),
                 ('__LECTURES__', json.dumps(lectures or [], ensure_ascii=False)),
                 ('__TOTAL__', str(total)),
                 ('__SLIDES__', '\n'.join(slides))]:
        out = out.replace(k, v)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--id', required=True)
    ap.add_argument('--name', default='')
    ap.add_argument('--out-dir', dest='out_dir', default=None)
    ap.add_argument('--classes-file', dest='classes_file', default=None)
    args = ap.parse_args()

    jpath = os.path.join(here, 'build', '%s.json' % args.id)
    deck = json.load(open(jpath, encoding='utf-8'))
    tpl = open(os.path.join(here, 'tools', 'deck_template.html'), encoding='utf-8').read()

    classes = []
    if args.classes_file:
        classes = json.load(open(args.classes_file, encoding='utf-8'))
    html_out = render(deck, tpl, classes=classes)
    fname = '%s-%s.html' % (args.id, args.name or '课件')
    outdir = args.out_dir or here
    if not os.path.isdir(outdir):
        os.makedirs(outdir)
    path = os.path.join(outdir, fname)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(html_out)
    print('HTML ->', path, '%.1f KB' % (os.path.getsize(path) / 1024))
    print('pages:', len(deck['slides']), '| 残留占位符:',
          re.findall(r'__[A-Z_]+__', html_out) or '无')


if __name__ == '__main__':
    main()
