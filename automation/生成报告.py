# -*- coding: utf-8 -*-
"""自动生成组合投资月度 HTML 报告（多文件版 + 自包含单文件版）

输入：
  投资组合分析明细.xlsx        （更新测算.py 产出：权重 / 指标 / 回测 / 回撤）
  单基金止损回测结果.xlsx      （止损回测.py 产出：指标对比_冷静期 / 每日明细_冷静20日）
  止损监控状态.json / 止损监控日志.csv / 持仓配置.json
  每日基金净值_自动_*.xlsx     （最新一个，取“区间回报”还原净值指数）
模板：
  报告模板.html                （含 {{XXX}} 占位标记）
输出目录：
  portfolio-report-YYYYMM/     （YYYYMM 取报告数据最新日期所在月）
    portfolio-report-YYYYMM.html
    assets/data.js  assets/charts.js  _shared/js/echarts.min.js
    portfolio-report-YYYYMM-单文件.html（内联 echarts + data.js + charts.js）

脚本可重复运行（幂等）；某 sheet 缺失时打印警告并把对应位置留空（—），不中断。
"""
import os
import re
import glob
import json
import math
import shutil

import numpy as np
import pandas as pd

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

PLAN = '投资组合分析明细.xlsx'
STOP = '单基金止损回测结果.xlsx'
STATE = '止损监控状态.json'
LOG = '止损监控日志.csv'
CFG = '持仓配置.json'
TEMPLATE = '报告模板.html'
DASH = '—'
COOLDOWN = 20
BOND = '华泰保兴安悦债券C'      # 债券基金：回撤 ≥3% 清仓转现金
BOND_DD = 0.03
LEVELS = [(0.10, 0.70), (0.18, 0.40), (0.25, 0.00)]   # 权益阶梯减仓：(回撤阈值, 目标仓位)

# 资产池 → 大类（顺序与报告附录 E 一致）
CATEGORY_GROUPS = [
    ('避险 / 商品', True, ['工银黄金ETF联接C']),
    ('固收', False, ['华泰保兴安悦债券C']),
    ('港股 / 红利', False, ['易方达恒生港股通红利低波动ETF联接发起式C']),
    ('宽基 / 量化', False, ['易方达中证1000量化增强C']),
    ('科技成长', False, ['嘉实上证科创板芯片ETF发起联接C', '华夏中证动漫游戏ETF发起式联接C',
                        '招商中证机器人ETF发起联接C', '东方新能源汽车主题混合']),
    ('医药消费', False, ['永赢医药创新智选混合发起C', '广发中证全指家用电器ETF联接C',
                        '招商中证白酒指数C', '华宝中证沪港深新消费指数C']),
    ('周期资源', False, ['富国中证煤炭指数C', '嘉实中证稀有金属主题ETF发起联接C',
                        '华泰柏瑞中证稀土产业ETF发起式联接C', '天弘中证细分化工产业主题ETF联接C']),
    ('金融 / 军工 / 农业', False, ['南方中证全指证券公司ETF联接C', '富国中证军工指数(LOF)C',
                                '农银现代农业加混合']),
]
FUND_CAT = {f: (cat, gold) for cat, gold, fs in CATEGORY_GROUPS for f in fs}

ROW = '          '   # tbody 内行缩进


# ---------------------------------------------------------------- 读取工具
def read_sheet(path, sheet, **kw):
    try:
        return pd.read_excel(path, sheet_name=sheet, **kw)
    except Exception as e:
        print(f'[警告] 无法读取 {os.path.basename(path)} 的 sheet「{sheet}」：{e}')
        return None


def read_weights(sheet):
    df = read_sheet(PLAN, sheet)
    if df is None or df.shape[1] < 2:
        print(f'[警告] 权重表「{sheet}」缺失，按空处理')
        return {}
    df = df.rename(columns={df.columns[0]: '基金', df.columns[1]: '权重'})
    df = df.dropna(subset=['基金'])
    out = {}
    for _, r in df.iterrows():
        try:
            out[str(r['基金'])] = float(r['权重'])
        except (TypeError, ValueError):
            continue
    return out


def read_overview(sheet):
    df = read_sheet(PLAN, sheet)
    if df is None or df.shape[1] < 2:
        print(f'[警告] 指标表「{sheet}」缺失，相关数值留空')
        return {}
    return {str(k): v for k, v in zip(df.iloc[:, 0], df.iloc[:, 1])}


def ok(v):
    if v is None:
        return False
    try:
        return not pd.isna(v)
    except (TypeError, ValueError):
        return False


def num(v):
    """把 '12.6964%' / '0.8559' 这类单元格转成小数"""
    if not ok(v):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    try:
        return float(s[:-1]) / 100 if s.endswith('%') else float(s)
    except ValueError:
        return None


def fp(v, nd=2, sign=False):
    if not ok(v):
        return DASH
    return f'{v * 100:+.{nd}f}%' if sign else f'{v * 100:.{nd}f}%'


def fn(v, nd=2):
    if not ok(v):
        return DASH
    return f'{v:.{nd}f}'


def spct(v, nd=2):
    """已经是百分数的值 → '+1.23%'"""
    return f'{v:+.{nd}f}%' if isinstance(v, (int, float)) else DASH


def fi(v):
    if not ok(v):
        return DASH
    return f'{int(round(v))}'


def parse_date(v):
    try:
        return pd.Timestamp(v).normalize()
    except Exception:
        return None


def block(rows):
    return ('\n' + ROW).join(rows)


# ---------------------------------------------------------------- 1. 基准口径
ov = read_overview('组合统计概览')
rp_ov = read_overview('风险平价统计概览')

as_of = parse_date(ov.get('数据结束日期'))
data_start = parse_date(ov.get('数据起始日期'))
if as_of is None:
    autos = sorted(glob.glob(os.path.join(BASE, '每日基金净值_自动_*.xlsx')))
    if autos:
        d = read_sheet(autos[-1], '区间回报', index_col=0)
        if d is not None:
            idx = pd.to_datetime(d.index, errors='coerce')
            as_of = idx[idx.notna()].max()
if as_of is None:
    as_of = pd.Timestamp.today().normalize()
    print('[警告] 无法确定数据截止日期，按运行当天处理')
if data_start is None:
    data_start = as_of
    print('[警告] 缺少数据起始日期，数据区间按截止日处理')

ym = as_of.strftime('%Y%m')
asof_s = as_of.strftime('%Y-%m-%d')

mv = read_weights('最优配置权重')
rp = read_weights('风险平价配置权重')
names = sorted(set(mv) | set(rp))
mv_nz = {k: v for k, v in mv.items() if v > 1e-4}
rp_nz = {k: v for k, v in rp.items() if v > 1e-4}
mv_count, rp_count, total = len(mv_nz), len(rp_nz), len(names)

# ---------------------------------------------------------------- 2. 组合指标
mv_ann = num(ov.get('组合年化收益率'))
mv_vol = num(ov.get('组合年化波动率'))
mv_sharpe = num(ov.get('夏普比率'))
rp_ann = num(rp_ov.get('组合年化收益率'))
rp_vol = num(rp_ov.get('组合年化波动率'))
rp_sharpe = num(rp_ov.get('夏普比率'))

dd_df = read_sheet(PLAN, '两模型回撤对比_近一年')
mv_dd = rp_dd = None
if dd_df is not None and {'均值方差_回撤', '风险平价_回撤'} <= set(dd_df.columns):
    mv_dd = float(pd.to_numeric(dd_df['均值方差_回撤'], errors='coerce').min())
    rp_dd = float(pd.to_numeric(dd_df['风险平价_回撤'], errors='coerce').min())
else:
    print('[警告] 回撤对比 sheet 缺失，最大回撤留空')

# 权重集中度
top4 = sum(sorted((v for v in mv.values() if v > 0), reverse=True)[:4])
max_rp_w = max(rp.values()) if rp else None
max_rp_fund = max(rp, key=rp.get) if rp else None
max_rp_label = '债券' if (max_rp_fund and '债券' in max_rp_fund) else (max_rp_fund or DASH)

# ---------------------------------------------------------------- 3. 区间回测
def first_last(sheet):
    """返回 (区间首日, 区间末日, 末值)"""
    d = read_sheet(PLAN, sheet)
    if d is None or d.empty or d.shape[1] < 2:
        print(f'[警告] 区间回测「{sheet}」缺失，相关数值留空')
        return None, None, None
    idx = pd.to_datetime(d.iloc[:, 0], errors='coerce')
    vals = pd.to_numeric(d.iloc[:, 1], errors='coerce')
    good = idx[idx.notna()]
    if good.empty:
        return None, None, None
    return good.iloc[0], good.iloc[-1], (float(vals.iloc[-1]) if ok(vals.iloc[-1]) else None)


w1_start, _, w1_mv = first_last('回测_最近一个月_起点0')
_, _, w3_mv = first_last('回测_最近三个月_起点0')
_, _, ytd_mv = first_last('回测_今年以来_起点0')
w1_rp_start, _, w1_rp = first_last('风险平价_回测_最近一个月_起点0')
w3_start, _, w3_rp = first_last('风险平价_回测_最近三个月_起点0')
ytd_start, _, ytd_rp = first_last('风险平价_回测_今年以来_起点0')
if w3_start is None:
    w3_start = first_last('回测_最近三个月_起点0')[0]
if ytd_start is None:
    ytd_start = first_last('回测_今年以来_起点0')[0]

# ---------------------------------------------------------------- 4. 止损回测
sc = read_sheet(STOP, '指标对比_冷静期')
base_m = stop_m = None
cd20_m = None
if sc is None:
    print('[警告] 止损回测指标缺失，第 3 章指标留空')
else:
    labels = sc.iloc[:, 0].astype(str)
    base_rows = sc[labels.str.strip() == '无止损']
    stop_rows_m = sc[labels.str.contains('冷静 20', na=False) | labels.str.contains('冷静20', na=False)]
    cmp_rows = sc[labels.str.contains('冷静 30', na=False) | labels.str.contains('冷静30', na=False)]
    base_m = base_rows.iloc[0].to_dict() if len(base_rows) else None
    stop_m = stop_rows_m.iloc[0].to_dict() if len(stop_rows_m) else None
    cmp_m = cmp_rows.iloc[0].to_dict() if len(cmp_rows) else None
    if stop_m is None:
        print('[警告] 未找到「阶梯止损 10%/18%/25% + 冷静 20 日」行，第 3 章指标留空')

bt = read_sheet(STOP, '每日明细_冷静20日')
bt_start = bt_days = None
if bt is not None and '日期' in bt.columns:
    dts = pd.to_datetime(bt['日期'], errors='coerce').dropna()
    if len(dts):
        bt_start, bt_days = dts.iloc[0], len(dts)
if bt_start is None:
    bt_start = ytd_start

bt_year = str(bt_start.year) if bt_start is not None else as_of.strftime('%Y')

# ---------------------------------------------------------------- 5. 1.2 止损监控
state = {}
if os.path.exists(STATE):
    try:
        with open(STATE, encoding='utf-8') as f:
            state = json.load(f)
    except Exception as e:
        print(f'[警告] {STATE} 读取失败：{e}')
funds = state.get('funds', {}) or {}

cfg = {}
if os.path.exists(CFG):
    try:
        with open(CFG, encoding='utf-8') as f:
            cfg = json.load(f)
    except Exception as e:
        print(f'[警告] {CFG} 读取失败：{e}')
capital = float(cfg.get('总资产', 0) or 0)
cash = float(cfg.get('现金', 0) or 0)
other_hold = [x for x in (cfg.get('其他持仓') or []) if isinstance(x, dict)]
base_dt = parse_date(cfg.get('基准日')) or as_of

prices = None
navs = {}
autos = sorted(glob.glob(os.path.join(BASE, '每日基金净值_自动_*.xlsx')))
if autos:
    d = read_sheet(autos[-1], '区间回报', index_col=0)
    if d is not None:
        d = d[~d.index.astype(str).str.contains('数据来源')]
        d.index = pd.to_datetime(d.index, errors='coerce')
        d = d[d.index.notna()].sort_index()
        d = d.dropna(axis=1, how='all').apply(pd.to_numeric, errors='coerce')
        prices = 1 + d.ffill() / 100.0
    nav = read_sheet(autos[-1], '单位净值', index_col=0)
    if nav is not None and not nav.empty:
        nav.index = pd.to_datetime(nav.index, errors='coerce')
        nav = nav[nav.index.notna()].sort_index()
        nav = nav.dropna(axis=1, how='all').apply(pd.to_numeric, errors='coerce')
        upto = nav.loc[:base_dt]
        if len(upto):
            navs = {str(k): float(v) for k, v in upto.iloc[-1].items() if ok(v)}
        else:
            print(f'[警告] 单位净值表无 {base_dt.strftime("%Y-%m-%d")} 及更早的数据，当前市值留空')
    else:
        print('[警告] 单位净值 sheet 缺失，当前市值留空')
else:
    print('[警告] 未找到自动净值文件，「当前净值指数」与回撤留空')

w_mv = {k: v for k, v in mv.items() if v > 1e-4}
w_sum = sum(w_mv.values())
if w_sum > 0:
    w_mv = {k: v / w_sum for k, v in w_mv.items()}

monitor_rows, flags = [], []
monitor_since = []
for c in sorted(w_mv, key=lambda x: -w_mv[x]):
    st = funds.get(c, {}) or {}
    peak = st.get('peak')
    p = None
    if prices is not None and c in prices.columns:
        s = prices[c].dropna()
        if len(s):
            p = float(s.iloc[-1])
    if not ok(p):
        monitor_rows.append(
            f'<tr><td>{c}</td><td class="num">{fp(w_mv[c])}</td><td class="num">{DASH}</td>'
            f'<td class="num">{DASH}</td><td>持有</td><td>{DASH}</td></tr>')
        continue
    if not ok(peak):
        peak = p
    dd = p / float(peak) - 1
    st_pos = float(st.get('pos', 1.0 if st.get('on', True) else 0.0))
    status = '持有' if st_pos >= 1 else (f'{st_pos:.0%}仓' if st_pos > 0 else '已止损')
    action = DASH
    want = 1.0
    if c == BOND:
        want = 0.0 if dd <= -BOND_DD else 1.0
    else:
        for thr, frac in LEVELS:
            if dd <= -thr + 1e-12:
                want = min(want, frac)
    if st_pos > want + 1e-9:
        if want <= 1e-9:
            action = '⚠ 止损清仓 → 转现金' if c == BOND else '⚠ 止损清仓 → 转债券'
            flags.append((c, '清仓', dd))
        else:
            action = f'⚠ 减至 {want:.0%} → 转债券'
            flags.append((c, f'减至{want:.0%}', dd))
    elif st_pos < 1:
        s = prices[c].dropna()
        pos = prices.index.get_loc(s.index[-1])
        off_i = None
        if st.get('last_off'):
            try:
                off_i = prices.index.get_loc(pd.Timestamp(st['last_off']))
            except KeyError:
                off_i = None
        cooled = True if off_i is None else (pos - off_i) >= COOLDOWN
        r20 = p / float(s.iloc[-21]) - 1 if len(s) > 20 else None
        if not cooled:
            action = f'冷静期中(还需{COOLDOWN - (pos - off_i)}日)'
        elif (r20 is not None and r20 > 0) and dd > -0.05:
            action = '✔ 满足回补条件 → 回补'
            flags.append((c, '回补', dd))
        else:
            action = '等待回补'
    if st.get('since'):
        monitor_since.append(str(st['since']))
    monitor_rows.append(
        f'<tr><td>{c}</td><td class="num">{fp(w_mv[c])}</td><td class="num">{p:.4f}</td>'
        f'<td class="num">{fp(dd)}</td><td>{status}</td><td>{action}</td></tr>')

# 当日动作：优先取监控日志（与 每日止损监控.py 落盘口径一致），否则用状态推导
today_actions = []
if os.path.exists(LOG):
    try:
        lg = pd.read_csv(LOG, encoding='utf-8-sig')
        if '日期' in lg.columns:
            sel = lg[lg['日期'].astype(str).str[:10] == asof_s]
            today_actions = [(str(r['基金']), str(r['事件']), r.get('自高点回撤')) for _, r in sel.iterrows()]
    except Exception as e:
        print(f'[警告] {LOG} 读取失败：{e}')
if not today_actions:
    today_actions = flags

anchor = max(monitor_since) if monitor_since else asof_s
if today_actions:
    parts = []
    for c, ev, d in today_actions:
        try:
            parts.append(f'{c}（{ev}，自高点回撤 {float(d) * 100:.2f}%）')
        except (TypeError, ValueError):
            parts.append(f'{c}（{ev}）')
    action_title = f'今日动作：触发 {len(today_actions)} 项'
    action_note = ('；'.join(parts) + '。触发后请按第 3 章规则处理：权益回撤 ≥10% 减至 70%、≥18% 减至 40%、≥25% 清仓'
                   '（资金转入债券，债券已清仓则转现金）；债券 ≥3% 清仓（资金转现金）→ '
                   '冷静 20 个交易日 → 双条件满足才回补。')
else:
    action_title = '今日动作：无'
    action_note = ('全部持仓未触发止损线（权益 10%/18%/25% 阶梯减至 70%/40%/清仓，债券 3% 清仓），'
                   '也没有处于冷静期或满足回补条件的基金。'
                   '触发后请按第 3 章规则处理：权益 ≥10% 减至 70%、≥18% 减至 40%、≥25% 清仓（资金转入债券，'
                   '债券已清仓则转现金）；债券 ≥3% 清仓转现金 → 冷静 20 个交易日 → '
                   '双条件满足才回补。')

# ---------------------------------------------------------------- 6. 表格行
def money(v, sign=False):
    if not ok(v):
        return DASH
    if abs(v) < 0.005:
        return '0.00'
    return f'{v:+,.2f}' if sign else f'{v:,.2f}'


def holding_value(c):
    """当前市值 = 真实份额 × 基准日单位净值；无持仓返回 0"""
    st = funds.get(c, {}) or {}
    sh = st.get('shares')
    if not ok(sh) or float(sh) <= 0:
        return 0.0
    nav = navs.get(c)
    if not ok(nav):
        return None
    return float(sh) * float(nav)


weight_rows = []
adj_buy, adj_sell = 0.0, 0.0
model_val = 0.0
for c in sorted(names, key=lambda x: (-mv.get(x, 0), -rp.get(x, 0), x)):
    cat, gold = FUND_CAT.get(c, ('其他', False))
    badge = f'<span class="badge gold">{cat}</span>' if gold else f'<span class="badge">{cat}</span>'
    cur = holding_value(c)
    adj = (mv.get(c, 0) * capital - cur) if ok(cur) else None
    if ok(adj):
        if adj > 0:
            adj_buy += adj
        elif adj < 0:
            adj_sell += -adj
    if ok(cur):
        model_val += cur
    adj_cls = 'pos' if ok(adj) and adj > 0 else ('neg' if ok(adj) and adj < 0 else '')
    weight_rows.append(
        f'<tr><td>{c}</td><td>{badge}</td><td class="num">{fp(mv.get(c, 0))}</td>'
        f'<td class="num">{money(cur)}</td>'
        f'<td class="num {adj_cls}">{money(adj, sign=True)}</td></tr>')

# 现金与其他持仓：未纳入模型池，权重记 0；计入市值合计
extra_val = 0.0
if cash > 0:
    extra_val += cash
    weight_rows.append(
        f'<tr><td>现金</td><td><span class="badge">现金</span></td>'
        f'<td class="num">{fp(0)}</td><td class="num">{money(cash)}</td>'
        f'<td class="num">{money(0)}</td></tr>')
for h in other_hold:
    v = float(h.get('市值', 0) or 0)
    extra_val += v
    nm = f'{h.get("代码", "")} {h.get("名称", "")}'.strip()
    weight_rows.append(
        f'<tr><td>{nm}</td><td><span class="badge">其他</span></td>'
        f'<td class="num">{fp(0)}</td><td class="num">{money(v)}</td>'
        f'<td class="num">{money(0)}</td></tr>')

weight_rows.append(
    f'<tr class="sum"><td>合计（模型 {len(names)} 只 + 现金 + 其他）</td><td>{DASH}</td>'
    f'<td class="num">{fp(sum(mv.values()))}</td>'
    f'<td class="num">{money(model_val + extra_val)}</td>'
    f'<td class="num {"pos" if adj_buy - adj_sell >= 0 else "neg"}">{money(adj_buy - adj_sell, sign=True)}</td></tr>')

other_text = '、'.join(
    f'{h.get("代码", "")} {h.get("名称", "")} {float(h.get("市值", 0) or 0):,.2f} 元'.strip()
    for h in other_hold) or '无其他持仓'


metric_rows = [
    f'<tr><td>持仓数量</td><td class="num">{mv_count} / {total}</td><td class="num">{rp_count} / {total}</td></tr>',
    f'<tr><td>权重集中度</td><td class="num">前 4 只合计 {top4 * 100:.1f}%</td>'
    f'<td class="num">最大 {fp(max_rp_w)}（{max_rp_label}）</td></tr>',
]


def wrow(label, start, m, p, note):
    cls_m = 'pos' if (ok(m) and m >= 0) else 'neg'
    cls_p = 'pos' if (ok(p) and p >= 0) else 'neg'
    return (f'<tr><td>{label}（{start.strftime("%Y-%m-%d") if start is not None else DASH} 起）</td>'
            f'<td class="num {cls_m}">{fp(m, sign=True)}</td>'
            f'<td class="num {cls_p}">{fp(p, sign=True)}</td><td>{note}</td></tr>')


window_rows = [
    wrow('近 1 个月', w1_start, w1_mv, w1_rp, 'RP 几乎持平'),
    wrow('近 3 个月', w3_start, w3_mv, w3_rp, '两个模型跌幅相近'),
    wrow('今年以来', ytd_start, ytd_mv, ytd_rp, 'MV 靠港股红利与债券扛住回撤'),
]

def stop_row(name, m, strong=False, neg_dd=False):
    if m is None:
        return f'<tr><td>{name}</td>' + ''.join(f'<td class="num">{DASH}</td>' for _ in range(6)) + '</tr>'
    ret_cls = 'pos' if (ok(m.get('累计收益')) and m['累计收益'] >= 0) else 'neg'
    sh_cls = 'pos' if (ok(m.get('夏普')) and m['夏普'] >= 0) else 'neg'
    whip_zero = ok(m.get('打脸次数(≤10日回补)')) and m['打脸次数(≤10日回补)'] == 0
    whip_attr = 'class="num pos"' if whip_zero else 'class="num"'
    dd_cls = 'class="num neg"' if neg_dd else 'class="num"'
    wrap = (lambda s: f'<strong>{s}</strong>') if strong else (lambda s: s)
    cells = [
        wrap(name),
        f'<td class="num {ret_cls}">{wrap(fp(m.get("累计收益"), sign=True))}</td>',
        f'<td class="num">{wrap(fp(m.get("年化波动")))}</td>',
        f'<td class="num {sh_cls}">{wrap(fn(m.get("夏普")))}</td>',
        f'<td {dd_cls}>{wrap(fp(m.get("最大回撤")))}</td>',
        f'<td class="num">{wrap(fi(m.get("止损事件数")))}</td>',
        f'<td {whip_attr}>{wrap(fi(m.get("打脸次数(≤10日回补)")))}</td>',
    ]
    return '<tr><td>' + cells[0] + '</td>' + ''.join(cells[1:]) + '</tr>'


stop_rows = [
    stop_row('无止损', base_m, neg_dd=True),
    stop_row('阶梯止损 10%/18%/25% + 冷静 20 日', stop_m, strong=True),
]

category_rows = []
for cat, gold, fs in CATEGORY_GROUPS:
    badge = f'<span class="badge gold">{cat}</span>' if gold else f'<span class="badge">{cat}</span>'
    category_rows.append(
        f'<tr><td>{badge}</td><td>{" / ".join(fs)}</td>'
        f'<td class="num">{fp(sum(v for k, v in mv.items() if k in fs))}</td>'
        f'<td class="num">{fp(sum(v for k, v in rp.items() if k in fs))}</td></tr>')

# ---------------------------------------------------------------- 6.5 相关系数矩阵（近一年日收益）
INDEX_CACHE = '指数行情.csv'
INDEXES = [('上证指数', '1.000001'), ('沪深300', '1.000300'),
           ('中证1000', '1.000852'), ('创业板指', '0.399006')]


def fetch_index_closes(secid, beg='20260101'):
    """东方财富日 K 线接口，返回 (日期 → 收盘价)"""
    import requests
    r = requests.get(
        'https://push2his.eastmoney.com/api/qt/stock/kline/get',
        params={'secid': secid, 'fields1': 'f1,f2,f3,f4,f5',
                'fields2': 'f51,f53', 'klt': '101', 'fqt': '0',
                'beg': beg, 'end': '20500101'},
        headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
                 'Referer': 'https://quote.eastmoney.com/'},
        timeout=20)
    r.raise_for_status()
    klines = (r.json().get('data') or {}).get('klines') or []
    rows = {}
    for line in klines:
        p = line.split(',')
        if len(p) >= 2:
            try:
                rows[pd.Timestamp(p[0])] = float(p[1])
            except ValueError:
                continue
    if not rows:
        raise RuntimeError('接口返回空数据')
    return pd.Series(rows).sort_index()


def load_index_data(as_of):
    """优先复用 指数行情.csv；缓存未覆盖截止日时联网抓取并回写"""
    cols = [n for n, _ in INDEXES]
    path = os.path.join(BASE, INDEX_CACHE)
    cached = None
    if os.path.exists(path):
        try:
            c = pd.read_csv(path, encoding='utf-8-sig')
            c['日期'] = pd.to_datetime(c['日期'], errors='coerce')
            c = c.dropna(subset=['日期']).set_index('日期').sort_index()
            cached = c.apply(pd.to_numeric, errors='coerce')
        except Exception as e:
            print(f'[警告] {INDEX_CACHE} 读取失败：{e}')
    if (cached is not None and all(n in cached.columns for n in cols)
            and cached.index.max() >= as_of):
        return cached[cols].loc[:as_of], f'本地缓存 {INDEX_CACHE}'
    data = {}
    for n, sid in INDEXES:
        try:
            data[n] = fetch_index_closes(sid)
        except Exception as e:
            print(f'[警告] 指数「{n}」抓取失败：{e}')
            if cached is not None and n in cached.columns:
                data[n] = cached[n].dropna()
    if not data:
        print('[警告] 无指数行情，图 7 数据为空')
        return None, ''
    df = pd.DataFrame(data).sort_index()
    df.index.name = '日期'
    try:
        df.to_csv(path, encoding='utf-8-sig')
        print(f'已更新指数行情缓存：{INDEX_CACHE}（{len(df)} 条）')
    except Exception as e:
        print(f'[警告] {INDEX_CACHE} 写入失败：{e}')
    return df[[n for n in cols if n in df.columns]].loc[:as_of], f'东方财富 K 线接口（已缓存 {INDEX_CACHE}）'


held_list = [k for k in funds
             if ok((funds.get(k) or {}).get('shares')) and float(funds[k]['shares']) > 0]

CORR_METHOD = 'Pearson'
FUND_COLS = list(prices.columns) if prices is not None else []
corr_order, corr_matrix, corr_sel_flags, corr_held_flags = [], [], [], []
corr_window, corr_avg_all, corr_avg_sel = DASH, DASH, DASH
corr_hi_pair = corr_lo_pair = DASH
corr_hi_val = corr_lo_val = DASH
corr_high_n, corr_obs = 0, 0
model_sel = [k for k in sorted(mv_nz, key=lambda x: -mv_nz[x]) if k in FUND_COLS]

if prices is not None and len(FUND_COLS) > 1:
    daily_ret = (prices.diff() / prices.shift(1)).dropna(how='all')
    r1y = daily_ret[daily_ret.index >= as_of - pd.DateOffset(years=1)].dropna(how='all')
    if len(r1y) > 20:
        cm = r1y.corr()
        corr_window = (f'{r1y.index.min().strftime("%Y-%m-%d")} ~ '
                       f'{r1y.index.max().strftime("%Y-%m-%d")}')
        corr_obs = len(r1y)
        rest = [k for k in cm.columns if k not in model_sel]
        rest.sort(key=lambda x: (-rp.get(x, 0), x))
        corr_order = model_sel + rest
        corr_matrix = [[round(float(cm.loc[a, b]), 3) if ok(cm.loc[a, b]) else None
                        for b in corr_order] for a in corr_order]
        corr_sel_flags = [1 if k in model_sel else 0 for k in corr_order]
        corr_held_flags = [1 if k in held_list else 0 for k in corr_order]
        mask = np.triu(np.ones(cm.shape), 1).astype(bool)
        pairs = cm.where(mask).stack().dropna().sort_values()
        if len(pairs):
            corr_hi_pair = f'{pairs.index[-1][0]} × {pairs.index[-1][1]}'
            corr_lo_pair = f'{pairs.index[0][0]} × {pairs.index[0][1]}'
            corr_hi_val = f'{float(pairs.iloc[-1]):.2f}'
            corr_lo_val = f'{float(pairs.iloc[0]):.2f}'
            corr_avg_all = f'{float(pairs.mean()):.2f}'
            corr_high_n = int((pairs > 0.8).sum())
        if len(model_sel) > 1:
            off = cm.loc[model_sel, model_sel].values[np.triu_indices(len(model_sel), 1)]
            off = off[~np.isnan(off)]
            corr_avg_sel = f'{float(off.mean()):.2f}' if len(off) else DASH
    else:
        print('[警告] 近一年日收益样本不足，相关系数矩阵留空')

model_only = [k for k in model_sel if k not in held_list]
held_only = [k for k in held_list if k not in model_sel]
diff_parts = []
if model_only:
    diff_parts.append('模型选中但当前未持有：' + '、'.join(
        f'{k}（模型权重 {mv[k] * 100:.2f}%，真实份额为 0）' for k in model_only))
if held_only:
    diff_parts.append('当前实际持有但未被模型选中：' + '、'.join(
        f'{k}（真实份额 {float(funds[k]["shares"]):,.2f}）' for k in held_only))
corr_legend_note = (
    f'★ 为 MV 模型选中（权重 &gt; 0，{len(model_sel)} 只）；◇ 为 '
    f'{base_dt.strftime("%Y-%m-%d")} 真实持有（{len(held_list)} 只）；'
    f'★◇ 表示两者一致。' + ('差异——' + '；'.join(diff_parts) + '。' if diff_parts else ''))

# ---------------------------------------------------------------- 6.6 实仓固定权重 vs 主要指数
ytd_start_dt = None
if prices is not None and len(prices):
    cand = prices.index[prices.index >= pd.Timestamp(f'{as_of.year}-01-01')]
    if len(cand):
        ytd_start_dt = cand[0]

idx_df, idx_source = load_index_data(as_of)
ytd_idx_js = {'dates': [], 'port': [], 'sh': [], 'hs300': [], 'zz1000': [], 'cyb': []}
port_ytd, ytd_excess = None, DASH
idx_ytd = {n: None for n, _ in INDEXES}
if prices is not None and idx_df is not None and ytd_start_dt is not None:
    vals = {}
    for k in held_list:
        if k in prices.columns and ok(navs.get(k)):
            vals[k] = float(funds[k]['shares']) * float(navs[k])
    missing_idx = [n for n, _ in INDEXES if n not in idx_df.columns]
    if missing_idx:
        print(f'[警告] 指数行情缺少 {missing_idx}，图 7 留空')
    elif vals:
        w_fix = {k: v / sum(vals.values()) for k, v in vals.items()}
        sub = prices.loc[prices.index >= ytd_start_dt, list(w_fix)].ffill()
        lvl = (sub * pd.Series(w_fix)).sum(axis=1)
        common = lvl.index.intersection(idx_df.index)
        common = common[common >= ytd_start_dt]
        idx_sub = idx_df.loc[common, [n for n, _ in INDEXES]].ffill()
        keep = idx_sub.notna().all(axis=1) & lvl.loc[common].notna()
        common = common[keep]
        if len(common) > 1:
            base_lvl = float(lvl.loc[common].iloc[0])
            base_idx = {n: float(idx_sub[n].iloc[0]) for n, _ in INDEXES}
            p = (lvl.loc[common] / base_lvl - 1) * 100
            idx_norm = {n: (idx_sub[n] / base_idx[n] - 1) * 100 for n, _ in INDEXES}
            ytd_idx_js = {
                'dates': [d.strftime('%Y-%m-%d') for d in common],
                'port': [round(float(v), 2) for v in p],
                'sh': [round(float(v), 2) for v in idx_norm['上证指数']],
                'hs300': [round(float(v), 2) for v in idx_norm['沪深300']],
                'zz1000': [round(float(v), 2) for v in idx_norm['中证1000']],
                'cyb': [round(float(v), 2) for v in idx_norm['创业板指']],
            }
            port_ytd = float(p.iloc[-1])
            idx_ytd = {n: float(idx_norm[n].iloc[-1]) for n, _ in INDEXES}
            ytd_excess = port_ytd - idx_ytd['沪深300']
        else:
            print('[警告] 组合与指数无共同交易日，图 7 留空')

# ---------------------------------------------------------------- 7. data.js
def series_sheet(sheet, col=None, name=None):
    d = read_sheet(PLAN, sheet, index_col=0)
    if d is None or d.empty:
        return None
    d.index = pd.to_datetime(d.index, errors='coerce')
    d = d[d.index.notna()].sort_index()
    s = d.iloc[:, 0] if col is None else d[col]
    s = pd.to_numeric(s, errors='coerce').dropna()
    s.name = name
    return s


ytd_a = series_sheet('回测_今年以来_起点0')
ytd_b = series_sheet('风险平价_回测_今年以来_起点0')
if ytd_a is not None and ytd_b is not None:
    j = pd.concat([ytd_a, ytd_b], axis=1, join='inner').dropna()
else:
    print('[警告] 今年以来回测缺失，图 2 数据为空')
    j = pd.DataFrame(columns=['mv', 'rp'])
ytd_js = {
    'dates': [d.strftime('%Y-%m-%d') for d in j.index],
    'mv': [round(float(v) * 100, 2) for v in j.iloc[:, 0]] if len(j) else [],
    'rp': [round(float(v) * 100, 2) for v in j.iloc[:, 1]] if len(j) else [],
}

if dd_df is not None and {'日期', '均值方差_回撤', '风险平价_回撤'} <= set(dd_df.columns):
    dd_s = dd_df.copy()
    dd_s['日期'] = pd.to_datetime(dd_s['日期'], errors='coerce')
    dd_s = dd_s.dropna(subset=['日期']).sort_values('日期')
    dd_js = {
        'dates': [d.strftime('%Y-%m-%d') for d in dd_s['日期']],
        'mv': [round(float(v) * 100, 2) for v in dd_s['均值方差_回撤']],
        'rp': [round(float(v) * 100, 2) for v in dd_s['风险平价_回撤']],
    }
else:
    print('[警告] 两模型回撤对比缺失，图 4 数据为空')
    dd_js = {'dates': [], 'mv': [], 'rp': []}

if bt is not None and {'日期', '无止损_累计收益', '冷静20日_累计收益'} <= set(bt.columns):
    bt_s = bt.copy()
    bt_s['日期'] = pd.to_datetime(bt_s['日期'], errors='coerce')
    bt_s = bt_s.dropna(subset=['日期']).sort_values('日期')
    stop_js = {
        'dates': [d.strftime('%Y-%m-%d') for d in bt_s['日期']],
        'base': [round(float(v) * 100, 2) for v in bt_s['无止损_累计收益']],
        'stop': [round(float(v) * 100, 2) for v in bt_s['冷静20日_累计收益']],
    }
else:
    print('[警告] 止损每日明细缺失，图 5 数据为空')
    stop_js = {'dates': [], 'base': [], 'stop': []}

payload = {
    'asOf': asof_s,
    'names': names,
    'mvW': [round(float(mv.get(n, 0)) * 100, 2) for n in names],
    'rpW': [round(float(rp.get(n, 0)) * 100, 2) for n in names],
    'ytd': ytd_js,
    'dd': dd_js,
    'stop': stop_js,
    'corr': {'names': corr_order, 'sel': corr_sel_flags, 'held': corr_held_flags,
             'matrix': corr_matrix, 'method': CORR_METHOD,
             'window': corr_window, 'obs': corr_obs},
    'ytdIdx': ytd_idx_js,
    'windows': {'labels': ['近1个月', '近3个月', '今年以来'],
                'mv': [round(w1_mv * 100, 2) if ok(w1_mv) else None,
                       round(w3_mv * 100, 2) if ok(w3_mv) else None,
                       round(ytd_mv * 100, 2) if ok(ytd_mv) else None],
                'rp': [round(w1_rp * 100, 2) if ok(w1_rp) else None,
                       round(w3_rp * 100, 2) if ok(w3_rp) else None,
                       round(ytd_rp * 100, 2) if ok(ytd_rp) else None]},
}
data_js = 'window.RPT = ' + json.dumps(payload, ensure_ascii=False) + ';\n'

# ---------------------------------------------------------------- 8. 填充模板
base_dd = num(base_m.get('最大回撤')) if base_m else None
stop_dd = num(stop_m.get('最大回撤')) if stop_m else None
base_vol = num(base_m.get('年化波动')) if base_m else None
stop_vol = num(stop_m.get('年化波动')) if stop_m else None
base_ret = num(base_m.get('累计收益')) if base_m else None
stop_ret = num(stop_m.get('累计收益')) if stop_m else None
cmp_ret = num(cmp_m.get('累计收益')) if cmp_m else None
cmp_dd = num(cmp_m.get('最大回撤')) if cmp_m else None
main_n = stop_m.get('止损事件数') if stop_m else None
main_whip = stop_m.get('打脸次数(≤10日回补)') if stop_m else None

repl = {
    'TITLE_MONTH': as_of.strftime('%Y-%m'),
    'KICKER_MONTH': as_of.strftime('%Y 年 %m 月').replace(' 0', ' '),
    'DATA_RANGE': f'{data_start.strftime("%Y-%m-%d")} ~ {asof_s}',
    'ASOF': asof_s,
    'BT_YEAR': bt_year,
    'MV_COUNT': str(mv_count),
    'ZERO_COUNT': str(total - mv_count),
    'RP_COUNT': str(rp_count),
    'RP_RC_EACH': f'{100 / rp_count:.2f}%' if rp_count else DASH,
    'WEIGHTS_ROWS': block(weight_rows),
    'BASE_DATE': base_dt.strftime('%Y-%m-%d'),
    'TOTAL_ASSETS': f'{capital:,.2f}',
    'CASH': f'{cash:,.2f}',
    'OTHER_HOLD_TEXT': other_text,
    'ADJ_BUY': f'{adj_buy:,.2f}',
    'ADJ_SELL': f'{adj_sell:,.2f}',
    'ADJ_NET': f'{adj_buy - adj_sell:+,.2f}',
    'METRICS_ROWS': block(metric_rows),
    'MAX_RP_W': fp(max_rp_w),
    'ANCHOR_DATE': anchor,
    'MONITOR_ROWS': block(monitor_rows),
    'ACTION_TITLE': action_title,
    'ACTION_NOTE': action_note,
    'WINDOWS_ROWS': block(window_rows),
    'DD_MIN_MV': fp(mv_dd, 1),
    'DD_MIN_RP': fp(rp_dd, 1),
    'BT_START': bt_start.strftime('%Y-%m-%d') if bt_start is not None else DASH,
    'BT_DAYS': str(bt_days) if bt_days else DASH,
    'STOP_ROWS': block(stop_rows),
    'NOTE_BASE_DD': fp(base_dd, 1),
    'NOTE_STOP_DD': fp(stop_dd, 1),
    'NOTE_BASE_VOL': fp(base_vol, 1),
    'NOTE_STOP_VOL': fp(stop_vol, 1),
    'NOTE_BASE_RET': fp(base_ret, 1, sign=True),
    'NOTE_STOP_RET': fp(stop_ret, 1, sign=True),
    'CD_MAIN_RET': fp(stop_ret, sign=True),
    'CD_MAIN_DD': fp(stop_dd),
    'CD_CMP_RET': fp(cmp_ret, sign=True),
    'CD_CMP_DD': fp(cmp_dd),
    'CD_MAIN_N': fi(main_n),
    'CD_MAIN_WHIP': fi(main_whip),
    'CATEGORY_ROWS': block(category_rows),
    'CORR_METHOD': CORR_METHOD,
    'CORR_WINDOW': corr_window,
    'CORR_OBS': str(corr_obs) if corr_obs else DASH,
    'CORR_AVG_ALL': corr_avg_all,
    'CORR_AVG_SEL': corr_avg_sel,
    'CORR_HIGH_PAIR': corr_hi_pair,
    'CORR_HIGH_VAL': corr_hi_val,
    'CORR_LOW_PAIR': corr_lo_pair,
    'CORR_LOW_VAL': corr_lo_val,
    'CORR_HIGH_N': str(corr_high_n),
    'CORR_LEGEND_NOTE': corr_legend_note,
    'YTD_START': ytd_start_dt.strftime('%Y-%m-%d') if ytd_start_dt is not None else DASH,
    'YTD_YEAR': str(as_of.year),
    'YTD_DAYS': str(len(ytd_idx_js['dates'])) if ytd_idx_js['dates'] else DASH,
    'YTD_PORT': spct(port_ytd),
    'YTD_SH': spct(idx_ytd.get('上证指数')),
    'YTD_HS300': spct(idx_ytd.get('沪深300')),
    'YTD_ZZ1000': spct(idx_ytd.get('中证1000')),
    'YTD_CYB': spct(idx_ytd.get('创业板指')),
    'YTD_EXCESS': (f'{ytd_excess:+.2f} 个百分点'
                   if isinstance(ytd_excess, (int, float)) else DASH),
    'IDX_SOURCE': idx_source or DASH,
}

outdir = f'portfolio-report-{ym}'
assets = os.path.join(outdir, 'assets')
shared = os.path.join(outdir, '_shared', 'js')
os.makedirs(assets, exist_ok=True)
os.makedirs(shared, exist_ok=True)

with open(os.path.join(assets, 'data.js'), 'w', encoding='utf-8') as f:
    f.write(data_js)


def find_asset(rel):
    """在当前输出目录或任一历史报告目录中找资源文件"""
    cand = os.path.join(outdir, rel)
    if os.path.exists(cand):
        return cand
    for d in sorted(glob.glob(os.path.join(BASE, 'portfolio-report-*'))):
        p = os.path.join(d, rel)
        if os.path.exists(p):
            return p
    return None


for rel in ('assets/charts.js', '_shared/js/echarts.min.js'):
    dst = os.path.join(outdir, rel)
    src = find_asset(rel)
    if src is None:
        print(f'[警告] 未找到 {rel}，请从历史报告目录复制')
        continue
    if os.path.abspath(src) != os.path.abspath(dst):
        shutil.copy2(src, dst)
        print(f'已复制 {rel} ← {src}')

with open(TEMPLATE, encoding='utf-8') as f:
    tpl = f.read()

SCRIPTS_MULTI = ('<script src="./_shared/js/echarts.min.js"></script>\n'
                 '<script src="assets/data.js"></script>\n'
                 '<script src="assets/charts.js"></script>')

html = tpl.replace('{{SCRIPTS}}', SCRIPTS_MULTI)
for k, v in repl.items():
    html = html.replace('{{' + k + '}}', str(v))
left = re.findall(r'{{[A-Z0-9_]+}}', html)
if left:
    print('[警告] 模板存在未填充占位标记：' + ', '.join(sorted(set(left))))

html_path = os.path.join(outdir, f'portfolio-report-{ym}.html')
with open(html_path, 'w', encoding='utf-8') as f:
    f.write(html)


def inline(tag):
    p = os.path.join(outdir, tag)
    if not os.path.exists(p):
        print(f'[警告] 单文件版缺少 {tag}，跳过内联')
        return ''
    with open(p, encoding='utf-8') as f:
        return f.read().replace('</script>', '<\\/script>')


echarts_js = inline('_shared/js/echarts.min.js')
charts_js = inline('assets/charts.js')
single = tpl.replace('{{SCRIPTS}}',
                     '<script>\n' + echarts_js + '\n</script>\n'
                     '<script>\n' + data_js + '\n</script>\n'
                     '<script>\n' + charts_js + '\n</script>')
for k, v in repl.items():
    single = single.replace('{{' + k + '}}', str(v))
single_path = os.path.join(outdir, f'portfolio-report-{ym}-单文件.html')
with open(single_path, 'w', encoding='utf-8') as f:
    f.write(single)

print(f'\n报告已生成：{html_path}')
print(f'单文件版：{single_path}')
print(f'数据截止 {asof_s}｜权重 {mv_count}/{total}（MV）{rp_count}/{total}（RP）')
print(f'年化 {fp(mv_ann)}｜波动 {fp(mv_vol)}｜夏普 {fn(mv_sharpe, 3)}｜近一年回撤 {fp(mv_dd)}'
      f'｜RP 年化 {fp(rp_ann)}｜夏普 {fn(rp_sharpe, 3)}｜回撤 {fp(rp_dd)}')
print(f'止损回测：无止损 {fp(base_ret, sign=True)} / 回撤 {fp(base_dd)}；'
      f'阶梯止损 10%/18%/25%（冷静 20 日）{fp(stop_ret, sign=True)} / 回撤 {fp(stop_dd)}')
if capital:
    print(f'监控总资产 {capital:,.2f} 元（持仓配置.json，基准 {base_dt.strftime("%Y-%m-%d")}，'
          f'现金 {cash:,.2f} 元）')
    print(f'调仓金额：买入 {adj_buy:,.2f} 元 / 卖出 {adj_sell:,.2f} 元 / 净额 {adj_buy - adj_sell:+,.2f} 元')
print(f'相关矩阵 {len(corr_order)}×{len(corr_order)}（{CORR_METHOD}，{corr_window}，{corr_obs} 个日收益）'
      f'｜最高 {corr_hi_pair} {corr_hi_val}｜模型选中 {len(model_sel)} 只，实际持有 {len(held_list)} 只')
print(f'实仓固定权重 YTD {spct(port_ytd)}（自 {ytd_start_dt.strftime("%Y-%m-%d") if ytd_start_dt is not None else DASH}），'
      f'指数 {len(ytd_idx_js["dates"])} 个交易日'
      f'：上证 {spct(idx_ytd.get("上证指数"))}｜沪深300 {spct(idx_ytd.get("沪深300"))}'
      f'｜中证1000 {spct(idx_ytd.get("中证1000"))}｜创业板 {spct(idx_ytd.get("创业板指"))}')
