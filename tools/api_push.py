# -*- coding: utf-8 -*-
"""
不依赖 git 客户端，直接用 GitHub Git Data API 推送本地仓库（保留完整提交历史）。

背景：某些网络环境下 git 的 TLS 握手会被拦截（gnutls_handshake 失败），
但 HTTPS API 正常，此时用本脚本绕过。

用法：
    GITHUB_TOKEN=xxx python3 api_push.py <本地仓库路径> <owner>/<repo> [分支名]
"""
import base64
import json
import os
import subprocess
import sys
import time
import urllib.error
import urllib.request

TOKEN = os.environ.get('GITHUB_TOKEN') or open('/tmp/gh_token').read().strip()
API = 'https://api.github.com'


def api(method, path, payload=None, retry=3):
    for attempt in range(retry):
        data = json.dumps(payload).encode() if payload is not None else None
        req = urllib.request.Request(API + path, data=data, method=method)
        req.add_header('Authorization', 'Bearer ' + TOKEN)
        req.add_header('Accept', 'application/vnd.github+json')
        req.add_header('User-Agent', 'api-push')
        if data:
            req.add_header('Content-Type', 'application/json')
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                return json.loads(r.read().decode())
        except urllib.error.HTTPError as e:
            body = e.read().decode()
            if e.code in (502, 503, 500) and attempt < retry - 1:
                time.sleep(2)
                continue
            sys.exit('✗ %s %s -> %s\n%s' % (method, path, e.code, body[:500]))
    sys.exit('✗ 重试耗尽: ' + path)


def git(repo, *args):
    out = subprocess.run(['git', '-C', repo] + list(args),
                         capture_output=True, check=True)
    return out.stdout


def main():
    repo, full, branch = sys.argv[1], sys.argv[2], (sys.argv[3] if len(sys.argv) > 3 else 'main')
    owner, name = full.split('/')

    commits = git(repo, 'rev-list', '--reverse', 'HEAD').decode().split()
    print('准备推送 %d 次提交 -> %s (%s)' % (len(commits), full, branch))

    blob_map, parent, n_blob = {}, None, 0
    for i, sha in enumerate(commits, 1):
        msg = git(repo, 'log', '-1', '--format=%B', sha).decode().strip()
        author = git(repo, 'log', '-1', '--format=%an%x00%ae%x00%aI', sha).decode().strip('\n')
        committer = git(repo, 'log', '-1', '--format=%cn%x00%ce%x00%cI', sha).decode().strip('\n')
        an, ae, ad = author.split('\0')
        cn, ce, cd = committer.split('\0')

        entries = []
        for line in git(repo, 'ls-tree', '-r', sha).decode().splitlines():
            meta, path = line.split('\t', 1)
            mode, _type, bsha = meta.split()
            if bsha not in blob_map:
                raw = git(repo, 'cat-file', 'blob', bsha)
                r = api('POST', '/repos/%s/%s/git/blobs' % (owner, name), {
                    'content': base64.b64encode(raw).decode(), 'encoding': 'base64'})
                blob_map[bsha] = r['sha']
                n_blob += 1
            entries.append({'path': path, 'mode': mode, 'type': 'blob', 'sha': blob_map[bsha]})

        tree = api('POST', '/repos/%s/%s/git/trees' % (owner, name), {'tree': entries})
        payload = {'message': msg, 'tree': tree['sha'],
                   'author': {'name': an, 'email': ae, 'date': ad},
                   'committer': {'name': cn, 'email': ce, 'date': cd}}
        if parent:
            payload['parents'] = [parent]
        c = api('POST', '/repos/%s/%s/git/commits' % (owner, name), payload)
        parent = c['sha']
        print('  [%d/%d] %s  %s' % (i, len(commits), c['sha'][:7], msg.splitlines()[0][:60]))

    # 空仓库没有 refs，用 POST 创建
    try:
        api('POST', '/repos/%s/%s/git/refs' % (owner, name),
            {'ref': 'refs/heads/' + branch, 'sha': parent})
    except SystemExit:
        api('PATCH', '/repos/%s/%s/git/refs/heads/%s' % (owner, name, branch),
            {'sha': parent, 'force': True})

    print('\n✓ 推送完成（新建 blob %d 个）' % n_blob)
    print('  https://github.com/%s' % full)


if __name__ == '__main__':
    main()
