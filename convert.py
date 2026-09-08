# -*- coding: utf-8 -*-
"""
PowerPoint -> 带「断点标注」的网页课件

把你已有的 .pptx 转成单文件网页：保留原排版，打开自动显示"这个班上次讲到第几页"，
下课时按一个键记录进度。同一份课件给多个班上课、各班进度不同步时特别有用。

两种用法
--------
1) 转单个文件（最快）
   python3 convert.py 第一讲.pptx --id L01 --name 导论 --title "第一讲·导论：叙事与我们"
   python3 convert.py 第一讲.pptx --classes "导演2601,导演2602,网新2601"

2) 用配置文件批量转（推荐，一次转完一整门课）
   cp config.example.json config.json   # 改里面的班级和讲次
   python3 convert.py --config config.json

生成的每个 HTML 都是单文件、离线可用，双击就能讲。
"""
import argparse
import io
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, 'tools'))

from pptx2html import convert as pptx_to_json   # noqa: E402
from render import render as render_html        # noqa: E402


def load_config(path):
    cfg = json.load(io.open(path, encoding='utf-8'))
    cfg.setdefault('course', '未命名课程')
    cfg.setdefault('course_id', re.sub(r'[^a-zA-Z0-9]', '', cfg['course'])[:24] or 'deck')
    cfg.setdefault('classes', [])
    cfg.setdefault('lectures', [])
    cfg.setdefault('pptx_dir', '.')
    cfg.setdefault('out_dir', './web')
    return cfg


def parse_classes(spec, count=0):
    """--classes "a,b,c" 或 --class-count 6 -> [{id,name}]"""
    if spec:
        names = [x.strip() for x in re.split(r'[,，;；]', spec) if x.strip()]
        return [{'id': 'c%d' % (i + 1), 'name': n} for i, n in enumerate(names)]
    return [{'id': 'c%d' % (i + 1), 'name': '班级%d' % (i + 1)} for i in range(count)]


def build_one(pptx_path, lid, name, title, cfg, out_dir, quiet=False):
    if not os.path.isfile(pptx_path):
        print('  ✗ 找不到文件：%s' % pptx_path)
        return None
    deck, diff = pptx_to_json(pptx_path, lid, title or name or lid)

    # 中间层 JSON（便于人工微调，也方便排查）
    build_dir = os.path.join(out_dir, '_build')
    if not os.path.isdir(build_dir):
        os.makedirs(build_dir)
    with io.open(os.path.join(build_dir, '%s.json' % lid), 'w', encoding='utf-8') as f:
        json.dump(deck, f, ensure_ascii=False, separators=(',', ':'))

    tpl = io.open(os.path.join(HERE, 'tools', 'deck_template.html'), encoding='utf-8').read()
    lectures = [{'id': l['id'], 'name': l.get('name', l['id']),
                 'file': '%s-%s.html' % (l['id'], l.get('name', l['id'])),
                 'title': l.get('title') or l.get('name', l['id'])}
                for l in cfg.get('lectures', [])]
    html = render_html(deck, tpl,
                       classes=cfg.get('classes', []),
                       lectures=lectures,
                       course=cfg.get('course', ''),
                       course_id=cfg.get('course_id', 'deck'))

    fname = '%s-%s.html' % (lid, name or lid)
    path = os.path.join(out_dir, fname)
    with io.open(path, 'w', encoding='utf-8') as f:
        f.write(html)

    left = re.findall(r'__[A-Z_]+__', html)
    status = '✓' if not left else '✗ 残留占位符 %s' % left
    print('  %s %-28s %2d 页  %.0f KB' % (status, fname, len(deck['slides']),
                                          os.path.getsize(path) / 1024))
    if diff and not quiet:
        print('      （%d 处需人工核对：%s）' % (len(diff),
              '、'.join('第%d页' % n for n, _ in diff[:6])))
    return path


def main():
    ap = argparse.ArgumentParser(
        description='把 pptx 转成带断点标注的网页课件',
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('pptx', nargs='?', help='要转换的 .pptx 文件')
    ap.add_argument('--config', help='配置文件（批量转换用）')
    ap.add_argument('--id', help='讲次编号，如 L01')
    ap.add_argument('--name', help='文件名后缀，如 导论')
    ap.add_argument('--title', help='课件标题（显示在大屏和浏览器标签上）')
    ap.add_argument('--course', default='未命名课程', help='课程名')
    ap.add_argument('--course-id', dest='course_id', help='课程标识（决定哪些课件共享进度）')
    ap.add_argument('--classes', help='班级名，逗号分隔："导演2601,导演2602"')
    ap.add_argument('--class-count', dest='class_count', type=int, default=8,
                    help='不指定班级名时，生成几个占位班级（默认 8）')
    ap.add_argument('--out-dir', dest='out_dir', help='输出目录')
    args = ap.parse_args()

    if args.config:
        cfg = load_config(args.config)
        out_dir = args.out_dir or cfg['out_dir'] or '.'
        out_dir = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(args.config)), out_dir))
        pptx_dir = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(args.config)),
                                                cfg['pptx_dir']))
        print('课程：%s' % cfg['course'])
        print('班级：%s' % ('、'.join(c['name'] for c in cfg['classes']) or '（未配置，可在网页里添加）'))
        print('输出：%s\n' % out_dir)
        if not cfg['lectures']:
            # 没列讲次就扫描目录
            found = sorted(f for f in os.listdir(pptx_dir) if f.lower().endswith('.pptx')) \
                if os.path.isdir(pptx_dir) else []
            for i, f in enumerate(found, 1):
                base = os.path.splitext(f)[0]
                cfg['lectures'].append({'id': 'L%02d' % i, 'name': base,
                                        'file': f, 'title': base})
            if not found:
                print('✗ 没找到 pptx，也没有在 config 里写 lectures')
                sys.exit(1)
        for l in cfg['lectures']:
            p = l.get('file') or ''
            if not p.lower().endswith('.pptx'):
                p += '.pptx'
            build_one(os.path.join(pptx_dir, p), l['id'], l.get('name', l['id']),
                      l.get('title'), cfg, out_dir)
        print('\n完成。打开输出目录里任意一个 HTML 即可。')
        return

    if not args.pptx:
        ap.print_help()
        sys.exit(1)

    cfg = {
        'course': args.course,
        'course_id': args.course_id or (re.sub(r'[^a-zA-Z0-9]', '', args.course)[:24] or 'deck'),
        'classes': parse_classes(args.classes, args.class_count),
        'lectures': [],
    }
    lid = args.id or 'L01'
    name = args.name or os.path.splitext(os.path.basename(args.pptx))[0]
    out_dir = os.path.abspath(args.out_dir or '.')
    print('课程：%s ｜ 班级：%s' % (cfg['course'],
                                    '、'.join(c['name'] for c in cfg['classes'])))
    build_one(args.pptx, lid, name, args.title, cfg, out_dir)
    print('\n完成：%s' % os.path.join(out_dir, '%s-%s.html' % (lid, name)))


if __name__ == '__main__':
    main()
