# -*- coding: utf-8 -*-
"""
pptx -> build/<ID>.json
把 PowerPoint 每页的形状、坐标、文本 run 样式、表格提取为中间层 JSON，
交给 render.py 渲染成单文件网页 PPT。

用法:
    python3 tools/pptx2html.py --in <file.pptx> --id L01 --title "导论"
"""
import argparse
import base64
import json
import os
import re

from pptx import Presentation
from pptx.util import Emu
from pptx.enum.dml import MSO_FILL_TYPE, MSO_THEME_COLOR
from pptx.oxml.ns import qn

# 画布基准：1280 x 720（对应 10in x 5.625in，1pt = 1280/720 px）
BASE_W = 1280
PT2PX = BASE_W / 720.0          # 1pt -> px
TABLE_STUB = 'table'


# ------------------------- 颜色 -------------------------
def read_theme_map(prs):
    """主题色 name -> hex。pptx 里用 theme color 时 python-pptx 拿不到 RGB，必须查 theme part。"""
    m = {}
    try:
        part = prs.part.package.parts
        theme_part = None
        for p in part:
            if 'theme' in p.partname and p.partname.endswith('.xml'):
                theme_part = p
                break
        if theme_part is None:
            return m
        xml = theme_part.blob
        root = _parse(xml)
        ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
        for child in root.iter():
            tag = child.tag.split('}')[-1]
            if tag == 'clrScheme':
                for c in child:
                    name = c.tag.split('}')[-1]
                    if name in ('dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2',
                                'accent3', 'accent4', 'accent5', 'accent6',
                                'hlink', 'folHlink'):
                        for v in c:
                            vtag = v.tag.split('}')[-1]
                            if vtag == 'srgbClr':
                                m[name] = '#' + v.get('val')
                            elif vtag == 'sysClr':
                                m[name] = '#' + (v.get('lastClr') or '000000')
                            break
                break
    except Exception:
        pass
    return m


def _parse(xml):
    from lxml import etree
    return etree.fromstring(xml)


def safe_color(color_obj, theme_map):
    """python-pptx ColorFormat -> '#RRGGBB' / None"""
    if color_obj is None:
        return None
    try:
        t = color_obj.type
    except Exception:
        return None
    if t is None:
        return None
    # RGB 直取
    try:
        rgb = color_obj.rgb
        if rgb is not None:
            return '#' + str(rgb)
    except Exception:
        pass
    # 主题色
    try:
        tc = color_obj.theme_color
        if tc is not None:
            key = {
                MSO_THEME_COLOR.DARK_1: 'dk1', MSO_THEME_COLOR.LIGHT_1: 'lt1',
                MSO_THEME_COLOR.DARK_2: 'dk2', MSO_THEME_COLOR.LIGHT_2: 'lt2',
                MSO_THEME_COLOR.ACCENT_1: 'accent1', MSO_THEME_COLOR.ACCENT_2: 'accent2',
                MSO_THEME_COLOR.ACCENT_3: 'accent3', MSO_THEME_COLOR.ACCENT_4: 'accent4',
                MSO_THEME_COLOR.ACCENT_5: 'accent5', MSO_THEME_COLOR.ACCENT_6: 'accent6',
                MSO_THEME_COLOR.HYPERLINK: 'hlink',
                MSO_THEME_COLOR.FOLLOWED_HYPERLINK: 'folHlink',
            }.get(tc)
            if key and key in theme_map:
                return theme_map[key]
    except Exception:
        pass
    return None


def shape_fill(sh, theme_map):
    """形状填充色 -> hex / None"""
    try:
        f = sh.fill
        if f.type == MSO_FILL_TYPE.SOLID:
            return safe_color(f.fore_color, theme_map)
        if f.type == MSO_FILL_TYPE.GRADIENT:
            # 渐变降级：取第一个 stop
            try:
                stops = f.gradient_stops
                for s in stops:
                    c = safe_color(s.color, theme_map)
                    if c:
                        return c
            except Exception:
                pass
    except Exception:
        pass
    return None


def gradient_flag(sh):
    try:
        return sh.fill.type == MSO_FILL_TYPE.GRADIENT
    except Exception:
        return False


def slide_bg(slide, theme_map):
    try:
        f = slide.background.fill
        if f.type == MSO_FILL_TYPE.SOLID:
            return safe_color(f.fore_color, theme_map)
    except Exception:
        pass
    return '#FFFFFF'


# ------------------------- 文本 -------------------------
ALIGN_MAP = {
    'CENTER': 'center', 'CENTER (2)': 'center',
    'RIGHT': 'right', 'LEFT': 'left', 'JUSTIFY': 'justify',
    None: 'left',
}


def ea_font(run_xml):
    """东亚字体在 rPr/a:ea/@typeface，python-pptx 的 font.name 只给 latin"""
    try:
        pr = run_xml.find(qn('a:rPr'))
        if pr is None:
            return None
        for tag in ('a:ea', 'a:cs'):
            node = pr.find(qn(tag))
            if node is not None:
                tf = node.get('typeface')
                if tf:
                    return tf
    except Exception:
        pass
    return None


def extract_runs(par):
    """段落 -> 带样式的 run 列表；段内手动换行用 \\n 表示"""
    out = []
    for r in par.runs:
        f = r.font
        size = None
        try:
            if f.size is not None:
                size = round(f.size.pt * PT2PX, 1)
        except Exception:
            pass
        ea = ea_font(r._r)
        # 中文内容优先用东亚字体
        latin = f.name or ''
        has_cjk = bool(re.search(r'[\u4e00-\u9fff]', r.text))
        if ea and has_cjk:
            family = ea
        else:
            family = ea or latin
        out.append({
            'tx': r.text,
            's': size,
            'b': bool(f.bold) if f.bold is not None else False,
            'i': bool(f.italic) if f.italic is not None else False,
            'u': bool(f.underline) if f.underline is not None else False,
            'c': None,           # 颜色稍后补（需要 theme）
            'f': family,
        })
    return out


def extract_paragraphs(sh, theme_map):
    if not sh.has_text_frame:
        return [], None, None
    tf = sh.text_frame
    anchor = None
    try:
        anchor = tf.vertical_anchor.name if tf.vertical_anchor else None
    except Exception:
        pass
    paras = []
    total_h = 0.0
    for par in tf.paragraphs:
        runs = extract_runs(par)
        for rr, orig in zip(runs, par.runs):
            rr['c'] = safe_color(orig.font.color, theme_map)
        align = ALIGN_MAP.get(str(par.alignment), 'left')
        ls, lsu = None, None
        try:
            if isinstance(par.line_spacing, (float, int)):
                ls, lsu = round(float(par.line_spacing), 2), 'x'
            elif par.line_spacing is not None:
                ls, lsu = round(par.line_spacing.pt * PT2PX, 1), 'px'
        except Exception:
            pass
        sb = None
        try:
            if par.space_before is not None:
                sb = round((par.space_before.pt if hasattr(par.space_before, 'pt')
                            else par.space_before) * PT2PX, 1)
        except Exception:
            pass
        sa = None
        try:
            if par.space_after is not None:
                sa = round((par.space_after.pt if hasattr(par.space_after, 'pt')
                            else par.space_after) * PT2PX, 1)
        except Exception:
            pass
        paras.append({'runs': runs, 'al': align, 'lh': ls, 'lsu': lsu, 'sb': sb, 'sa': sa})

        # 估算高度，用于发现溢出
        max_sz = max([x['s'] for x in runs if x['s']] or [24])
        total_h += (max_sz or 24) * (ls if lsu == 'x' else 1.2)
    return paras, anchor, total_h


# ------------------------- 表格 -------------------------
def extract_table(sh, theme_map):
    tbl = sh.table
    rows = []
    col_w = [round(Emu(c.width) / 914400 * 128, 1) for c in tbl.columns]
    for row in tbl.rows:
        cells = []
        for cell in row.cells:
            paras = []
            for par in cell.text_frame.paragraphs:
                runs = extract_runs(par)
                for rr, orig in zip(runs, par.runs):
                    rr['c'] = safe_color(orig.font.color, theme_map)
                paras.append({'runs': runs,
                              'al': ALIGN_MAP.get(str(par.alignment), 'left'),
                              'lh': None, 'lsu': None, 'sb': None, 'sa': None})
            fill = None
            try:
                f = cell.fill
                if f.type == MSO_FILL_TYPE.SOLID:
                    fill = safe_color(f.fore_color, theme_map)
            except Exception:
                pass
            cells.append({'p': paras, 'fill': fill})
        rows.append({'h': round(Emu(row.height) / 914400 * 128, 1), 'cells': cells})
    return rows, col_w


# ------------------------- 主流程 -------------------------
def convert(pptx_path, deck_id, title, embed_images=True):
    prs = Presentation(pptx_path)
    theme_map = read_theme_map(prs)
    sw = prs.slide_width
    sh_h = prs.slide_height
    base_h = round(BASE_W * (sh_h / sw), 1)

    def px(emu):
        return round(Emu(emu) / sw * BASE_W, 1)

    slides = []
    diff = []

    for n, slide in enumerate(prs.slides, 1):
        bg = slide_bg(slide, theme_map)
        shapes = []
        page_notes = []

        for sh in slide.shapes:
            x, y, w, h = (px(sh.left), px(sh.top), px(sh.width), px(sh.height))
            fill = shape_fill(sh, theme_map)
            radius = 0
            try:
                if sh.auto_shape_type is not None and 'ROUNDED' in str(sh.auto_shape_type):
                    radius = 6
            except Exception:
                pass

            # 图片
            if str(sh.shape_type).startswith('PICTURE'):
                src = None
                try:
                    blob = sh.image.blob
                    ext = sh.image.ext
                    b64 = base64.b64encode(blob).decode()
                    src = 'data:image/%s;base64,%s' % (ext, b64)
                except Exception:
                    pass
                shapes.append({'t': 'i', 'x': x, 'y': y, 'w': w, 'h': h, 'src': src})
                continue

            # 表格
            if sh.has_table:
                rows, col_w = extract_table(sh, theme_map)
                shapes.append({'t': 'tb', 'x': x, 'y': y, 'w': w, 'h': h,
                               'rows': rows, 'colW': col_w})
                try:
                    _t = ''.join(r['tx'] for p in rows[0]['cells'][0]['p']
                                 for r in p['runs'])[:20]
                except Exception:
                    _t = ''
                page_notes.append('含表格：' + _t)
                continue

            # 文本框 / 矩形
            paras, anchor, est_h = extract_paragraphs(sh, theme_map)
            text = sh.text_frame.text if sh.has_text_frame else ''
            if gradient_flag(sh):
                page_notes.append('含渐变填充，已取首色：（%s）' % (fill or '无色'))
            # 溢出风险：估算文字高度明显超过框高（只检测有文字的框，跳过装饰色条）
            if paras and est_h and h and est_h > h * 1.35 and text.strip():
                page_notes.append('文字可能溢出框外（估算 %.0fpx / 框高 %.0fpx）：「%s」'
                                  % (est_h, h, text[:24].replace('\n', ' / ')))
            shapes.append({
                't': 't',
                'x': x, 'y': y, 'w': w, 'h': h,
                'fill': fill, 'radius': radius,
                'anchor': anchor,
                'p': paras,
            })

        if page_notes:
            diff.append((n, page_notes))
        slides.append({'n': n, 'bg': bg, 'shapes': shapes})

    deck = {
        'id': deck_id,
        'title': title,
        'size': {'w': BASE_W, 'h': base_h},
        'slides': slides,
    }
    return deck, diff


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--in', dest='src', required=True)
    ap.add_argument('--id', required=True)
    ap.add_argument('--title', default='')
    ap.add_argument('--out', default=None)
    args = ap.parse_args()

    here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    build = os.path.join(here, 'build')
    os.makedirs(build, exist_ok=True)

    deck, diff = convert(args.src, args.id, args.title)
    jpath = os.path.join(build, '%s.json' % args.id)
    with open(jpath, 'w', encoding='utf-8') as f:
        json.dump(deck, f, ensure_ascii=False, separators=(',', ':'))
    print('JSON ->', jpath, '%.1f KB' % (os.path.getsize(jpath) / 1024))

    dpath = os.path.join(build, '%s.diff.md' % args.id)
    lines = ['# %s 转换校对清单' % args.id + '\n',
             '自动转换不保证像素级一致。下列页面需要你打开 HTML 与原 PPT 对照看一眼：\n']
    if diff:
        for n, notes in diff:
            lines.append('\n## 第 %d 页\n' % n)
            lines += ['- %s' % s for s in notes]
    else:
        lines.append('\n未发现明显风险页。\n')
    with open(dpath, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    print('DIFF ->', dpath)


if __name__ == '__main__':
    main()
