# -*- coding: utf-8 -*-
"""统一推送模块：读取 推送配置.json，通过飞书群自定义机器人发送文本消息。

配置示例（推送配置.json）：
{
  "webhook": "https://open.feishu.cn/open-apis/bot/v2/hook/xxxx",
  "secret": ""
}
secret 仅在机器人安全设置选了"签名校验"时才需要填。
"""
import os, json, time, hmac, hashlib, base64
import requests

CFG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '推送配置.json')


def _cfg():
    """优先读环境变量（GitHub Actions 用 Secrets），其次读 推送配置.json。"""
    cfg = {}
    if os.path.exists(CFG_FILE):
        with open(CFG_FILE, encoding='utf-8') as f:
            cfg = json.load(f)
    env_url = (os.environ.get('FEISHU_WEBHOOK') or '').strip()
    env_secret = (os.environ.get('FEISHU_SECRET') or '').strip()
    if env_url:
        cfg['webhook'] = env_url
    if env_secret:
        cfg['secret'] = env_secret
    return cfg


def push_card(title, blocks, template='blue', quiet=False):
    """发送飞书卡片消息。blocks 为 lark_md 文本段落列表，元素间自动加分隔线。"""
    cfg = _cfg()
    url = (cfg.get('webhook') or '').strip()
    if not url:
        if not quiet:
            print('[推送] 未配置 webhook（推送配置.json），跳过')
        return False

    elements = []
    for i, b in enumerate(blocks):
        if i:
            elements.append({'tag': 'hr'})
        elements.append({'tag': 'div', 'text': {'tag': 'lark_md', 'content': b}})

    card = {
        'config': {'wide_screen_mode': True},
        'header': {'title': {'tag': 'plain_text', 'content': title}, 'template': template},
        'elements': elements,
    }
    body = {'msg_type': 'interactive', 'card': card}
    return _post(url, cfg, body, quiet)


def push_table_card(title, intro='', columns=None, rows=None, footer='',
                    template='blue', quiet=False):
    """发送飞书卡片 2.0，正文支持表格。

    columns: [(字段名, 表头显示名, 宽度)]，宽度用 'auto' 或像素字符串
    rows:    [{字段名: 值}]，值统一传字符串
    """
    cfg = _cfg()
    url = (cfg.get('webhook') or '').strip()
    if not url:
        if not quiet:
            print('[推送] 未配置 webhook（推送配置.json），跳过')
        return False

    elements = []
    if intro:
        elements.append({'tag': 'markdown', 'content': intro})
    if columns and rows:
        elements.append({
            'tag': 'table',
            'page_size': 10,
            'row_height': 'low',
            'header_style': {'text_align': 'left', 'text_size': 'normal',
                             'background_style': 'grey', 'bold': True, 'lines': 1},
            'columns': [{'name': n, 'display_name': d, 'data_type': 'text', 'width': w}
                        for n, d, w in columns],
            'rows': rows,
        })
    if footer:
        elements.append({'tag': 'markdown', 'content': footer})

    card = {
        'schema': '2.0',
        'config': {'update_multi': True},
        'header': {'title': {'tag': 'plain_text', 'content': title}, 'template': template},
        'body': {'direction': 'vertical', 'elements': elements},
    }
    return _post(url, cfg, {'msg_type': 'interactive', 'card': card}, quiet)


def _post(url, cfg, body, quiet=False):
    secret = (cfg.get('secret') or '').strip()
    if secret:
        ts = str(int(time.time()))
        string_to_sign = f'{ts}\n{secret}'
        body['timestamp'] = ts
        body['sign'] = base64.b64encode(
            hmac.new(string_to_sign.encode('utf-8'), digestmod=hashlib.sha256).digest()
        ).decode()
    try:
        r = requests.post(url, json=body, timeout=10)
        j = r.json()
        if r.status_code == 200 and j.get('code') == 0:
            if not quiet:
                print('[推送] 已发送')
            return True
        print(f'[推送] 失败：HTTP {r.status_code} {j}')
        return False
    except Exception as e:
        print(f'[推送] 异常：{e}')
        return False


def push(text, quiet=False):
    cfg = _cfg()
    url = (cfg.get('webhook') or '').strip()
    if not url:
        if not quiet:
            print('[推送] 未配置 webhook（推送配置.json），跳过')
        return False
    return _post(url, cfg, {'msg_type': 'text', 'content': {'text': text}}, quiet)
