# -*- coding: utf-8 -*-
"""每日止损监控（单基金口径，双击即可运行）

流程：
  1) 抓取最新净值（调用 基金数据获取.py 生成最新 每日基金净值_自动_*.xlsx）
  2) 读取本月配置权重（投资组合分析明细.xlsx 的“最优配置权重”）
  3) 逐只基金检查（仓位 pos ∈ 1.0 满仓 / 0.5 半仓 / 0.0 清仓）：
       权益类：自持有期高点回撤 ≥10% → 减半仓；≥15% → 清仓；
       债券类：回撤 ≥2% → 清仓，资金转现金（收益按 0 计）；
       权益止损腾出的资金转入债券基金；债券自身已清仓时转现金；
     回补条件：冷静期满10个交易日 且 近20日累计收益>0 且 回撤修复至 -5% 以内 → 回到满仓
  4) 输出今日状态表、触发清单，追加写入 止损监控日志.csv，状态保存在 止损监控状态.json

用法：
  python 每日止损监控.py            # 自动抓取最新数据后监控
  python 每日止损监控.py --no-fetch # 不联网，直接用本地最新的自动数据文件
"""
import os, sys, glob, json, subprocess
import pandas as pd
import numpy as np

base_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(base_dir)

BOND = '华泰保兴安悦债券C'
HALF_DD = 0.10        # 权益类：自高点回撤 10% 减半仓
FULL_DD = 0.15        # 权益类：自高点回撤 15% 清仓
BOND_DD = 0.03        # 债券基金：自高点回撤 3% 清仓（资金转现金）
RECOVER_DD = -0.05    # 回补条件之一：回撤修复至 -5% 以内
RECOVER_WIN = 20      # 回补条件之二：近 20 日累计收益 > 0
COOLDOWN = 10         # 冷静期：动作后 10 个交易日内禁止回补

STATE_FILE = os.path.join(base_dir, '止损监控状态.json')
LOG_FILE = os.path.join(base_dir, '止损监控日志.csv')
PLAN_FILE = os.path.join(base_dir, '投资组合分析明细.xlsx')

# ---------- 1. 抓取最新数据 ----------
if '--no-fetch' not in sys.argv:
    print('正在抓取最新净值……')
    r = subprocess.run([sys.executable, '基金数据获取.py'], capture_output=True, text=True,
                       encoding='utf-8', errors='ignore')
    if r.returncode != 0:
        print('抓取失败，改用本地已有数据：', r.stdout[-300:])

auto_files = sorted(glob.glob(os.path.join(base_dir, '每日基金净值_自动_*.xlsx')))
file_path = auto_files[-1] if auto_files else os.path.join(base_dir, '每日基金净值.xls')

df_raw = pd.read_excel(file_path, sheet_name='区间回报', index_col=0)
df_raw = df_raw[~df_raw.index.astype(str).str.contains('数据来源')]
df_raw.index = pd.to_datetime(df_raw.index, errors='coerce')
df_raw = df_raw[df_raw.index.notna()].sort_index()
df_raw = df_raw.dropna(axis=1, how='all').apply(pd.to_numeric, errors='coerce')
df_raw = df_raw.ffill().dropna()
df_cum = df_raw / 100.0
prices = 1 + df_cum
last_date = df_cum.index[-1]

# ---------- 2. 读取本月配置权重 ----------
plan = pd.read_excel(PLAN_FILE, sheet_name='最优配置权重')
plan.columns = ['基金', '权重']
W = plan.set_index('基金')['权重']
W = W[W > 1e-4]
W = W / W.sum()

plan_end = None
try:
    ov = pd.read_excel(PLAN_FILE, sheet_name='组合统计概览')
    row = ov[ov.iloc[:, 0].astype(str).str.contains('数据结束日期')]
    plan_end = str(pd.to_datetime(row.iloc[0, 1]).date())
except Exception:
    pass

# ---------- 2b. 单位净值与资金口径 ----------
CFG_FILE = os.path.join(base_dir, '持仓配置.json')
CAPITAL = 0.0
# 优先读环境变量（GitHub Actions Secret TOTAL_ASSET 注入），其次读本地配置
_env_capital = os.environ.get('TOTAL_ASSET', '').strip()
if _env_capital:
    try:
        CAPITAL = float(_env_capital)
    except ValueError:
        pass
if CAPITAL == 0.0 and os.path.exists(CFG_FILE):
    try:
        with open(CFG_FILE, 'r', encoding='utf-8') as f:
            CAPITAL = float(json.load(f).get('总资产', 0) or 0)
    except Exception:
        pass

NAV = None
try:
    NAV = pd.read_excel(file_path, sheet_name='单位净值', index_col=0)
    NAV.index = pd.to_datetime(NAV.index, errors='coerce')
    NAV = NAV[NAV.index.notna()].apply(pd.to_numeric, errors='coerce').sort_index().ffill()
except Exception as e:
    print(f'单位净值读取失败，金额列将留空：{e}')


def nav_at(c, d):
    if NAV is None or c not in NAV.columns:
        return np.nan
    s = NAV[c].dropna()
    if not len(s):
        return np.nan
    s = s[s.index <= pd.Timestamp(d)]
    return float(s.iloc[-1]) if len(s) else np.nan


def nav_prev(c, d):
    """d 之前最近一个交易日的单位净值（用于日涨跌/日收益）"""
    if NAV is None or c not in NAV.columns:
        return np.nan
    s = NAV[c].dropna()
    s = s[s.index < pd.Timestamp(d)]
    return float(s.iloc[-1]) if len(s) else np.nan


def nav_month_start(c, d):
    """上月最后一个交易日的单位净值（用于本月涨跌/本月收益）"""
    if NAV is None or c not in NAV.columns:
        return np.nan
    s = NAV[c].dropna()
    month_start = pd.Timestamp(d).replace(day=1)
    s = s[s.index < month_start]
    return float(s.iloc[-1]) if len(s) else np.nan


def peak_date(c):
    """自监控起点以来净值指数最高点出现的日期（回测口径）"""
    if c not in prices.columns:
        return '—'
    st = funds.get(c, {})
    since_str = st.get('since')
    try:
        since = pd.Timestamp(since_str) if since_str else prices.index[0]
    except Exception:
        since = prices.index[0]
    s = prices[c].dropna()
    s = s[s.index >= since]
    if not len(s):
        return '—'
    return str(s.idxmax().date())

print(f'数据文件：{os.path.basename(file_path)}    最新净值日：{last_date.date()}')
print(f'本月配置权重（来源：投资组合分析明细.xlsx，计划数据截止 {plan_end}）')
print('高点口径：自“监控起点”至今的最高净值，跨月连续跟踪，仅清仓后才重置')

# ---------- 3. 读取/初始化状态 ----------
state = {}
if os.path.exists(STATE_FILE):
    with open(STATE_FILE, 'r', encoding='utf-8') as f:
        state = json.load(f)

funds = state.setdefault('funds', {})

# 跨月连续跟踪：月度调仓只更换权重，不清空高点；仅当基金被调出组合（清仓）时才移除状态
# 调出组合但仍持有的基金：权重记 0 继续监控；只有份额为 0 才真正移除
held_not_in_w = [c for c in list(funds)
                 if c not in set(W.index) and float(funds[c].get('shares', 0) or 0) > 0]
for c in held_not_in_w:
    W[c] = 0.0
removed = [c for c in list(funds)
           if c not in set(W.index) and float(funds[c].get('shares', 0) or 0) <= 0]
for c in removed:
    funds.pop(c)
if removed:
    print(f'调仓剔除（已清仓，停止跟踪）：{"、".join(removed)}')
if held_not_in_w:
    print(f'调出组合但仍持有（权重 0，继续监控）：{"、".join(held_not_in_w)}')

# 监控起点：本轮建仓日（以最新净值日为起点）；之后跨月连续跟踪，不再重置
ANCHOR = last_date
for c in W.index:
    st = funds.setdefault(c, {})
    if 'peak' not in st or 'since' not in st:
        start = max(ANCHOR, prices[c].dropna().index[0])
        hist = prices.loc[start:, c].dropna()
        st['peak'] = float(hist.max()) if len(hist) else float(prices[c].dropna().iloc[-1])
        st.setdefault('on', True)
        st.setdefault('last_off', None)
        st['since'] = str(start.date())
    # 仓位字段：1.0 满仓 / 0.5 半仓 / 0.0 清仓；旧状态用 on 迁移
    st['pos'] = float(st.get('pos', 1.0 if st.get('on', True) else 0.0))
state['plan_date'] = plan_end

# 份额台账：以真实持仓为基准；状态中无份额记录的基金视为尚未持有（份额 0），由月度调仓建立
if CAPITAL > 0:
    for c in W.index:
        st = funds[c]
        if 'shares' not in st:
            st['shares'] = 0.0
    print(f'资金口径：总资产 {CAPITAL:,.2f} 元')

# ---------- 4. 逐只基金判定 ----------
# 第一遍：更新高点与仓位（不含文案，避免债券处理顺序影响资金去向判断）
info = {}
for c in W.index:
    st = funds[c]
    p = float(prices[c].dropna().iloc[-1])
    st['peak'] = max(st['peak'], p)
    dd = p / st['peak'] - 1
    bar_i = prices.index.get_loc(prices[c].dropna().index[-1])
    r20 = (p / float(prices[c].dropna().iloc[-RECOVER_WIN - 1]) - 1
           if len(prices[c].dropna()) > RECOVER_WIN else np.nan)
    stop_thr = BOND_DD if c == BOND else FULL_DD
    pos_q = float(st.get('pos', 1.0))
    act = None
    if pos_q >= 1.0:
        if dd <= -stop_thr:
            st['pos'] = 0.0
            st['on'] = False
            st['last_off'] = str(last_date.date())
            act = '清仓'
        elif c != BOND and dd <= -HALF_DD:
            st['pos'] = 0.5
            st['last_off'] = str(last_date.date())
            act = '减半仓'
    else:
        cooled = True
        if st.get('last_off'):
            off_i = prices.index.get_loc(pd.Timestamp(st['last_off']))
            cooled = (bar_i - off_i) >= COOLDOWN
        if not cooled:
            act = '冷静期'
        elif (r20 is not None and not np.isnan(r20) and r20 > 0) and dd > RECOVER_DD:
            st['pos'] = 1.0
            st['on'] = True
            st['last_off'] = None
            act = '回补'
    info[c] = {'p': p, 'dd': dd, 'bar_i': bar_i, 'r20': r20, 'act': act,
               'pos': float(st['pos'])}

bond_on = info.get(BOND, {}).get('pos', 1.0) > 0      # 债券是否仍在仓（决定止损资金去向）

# 第二遍：生成状态表与今日动作文案
rows, actions = [], []
for c in W.index:
    st, f = funds[c], info[c]
    p, dd, act, pos_q = f['p'], f['dd'], f['act'], f['pos']
    status = '持有' if pos_q >= 1.0 else ('半仓' if pos_q > 0 else '已止损')
    action = ''
    if act == '清仓':
        status = '已止损'
        to = '现金' if (c == BOND or not bond_on) else '债券'
        action = f'⚠ 止损清仓 → 转{to}'
        actions.append((c, '清仓', dd))
    elif act == '减半仓':
        status = '半仓'
        action = '⚠ 减半仓 → 转' + ('债券' if bond_on else '现金')
        actions.append((c, '减半仓', dd))
    elif act == '冷静期':
        off_i = prices.index.get_loc(pd.Timestamp(st['last_off']))
        action = f'冷静期中(还需{COOLDOWN - (f["bar_i"] - off_i)}日)'
    elif act == '回补':
        status = '持有'
        action = '✔ 满足回补条件 → 回补满仓'
        actions.append((c, '回补', dd))
    elif pos_q <= 0:
        action = '等待回补'
    elif pos_q < 1:
        action = '半仓持有，等待回补'

    rows.append({'基金': c, '权重': f'{W[c]:.2%}', '仓位': f'{pos_q:.0%}',
                 '监控起点': st.get('since', '—'), '当前净值指数': round(p, 4),
                 '自高点回撤': f'{dd:.2%}', '状态': status, '动作/提示': action})

res = pd.DataFrame(rows).sort_values('权重', key=lambda s: s.str.rstrip('%').astype(float), ascending=False)

print('\n===== 今日持仓监控 =====')
print(res.to_string(index=False))

stopped = [c for c in W.index if funds[c]['pos'] <= 0]
half = [c for c in W.index if 0 < funds[c]['pos'] < 1]
if stopped or half:
    moved = W[stopped].sum() + W[half].sum() * 0.5
    dest = f'{BOND}（债券当前实际权重 {W.get(BOND, 0) + (moved if BOND not in stopped else 0):.2%}）' \
        if bond_on else '现金（债券自身已清仓，止损资金不再买债券）'
    print(f'\n已清仓 {len(stopped)} 只、半仓 {len(half)} 只，腾出权重合计 {moved:.2%}，资金去向：{dest}')

print('\n===== 今日动作 =====')
if actions:
    for c, a, dd in actions:
        print(f'  {a}：{c}（自高点回撤 {dd:.2%}）')
else:
    print('  无（全部持仓未触发止损线，也没有满足回补条件的基金）')

# ---------- 5. 落盘 ----------
with open(STATE_FILE, 'w', encoding='utf-8') as f:
    json.dump(state, f, ensure_ascii=False, indent=2)

log = pd.DataFrame([{'日期': last_date.date(), '基金': c, '事件': a, '自高点回撤': round(dd, 4)}
                    for c, a, dd in actions],
                   columns=['日期', '基金', '事件', '自高点回撤'])
if os.path.exists(LOG_FILE) and len(log):
    log.to_csv(LOG_FILE, mode='a', header=False, index=False, encoding='utf-8-sig')
elif not os.path.exists(LOG_FILE):
    log.to_csv(LOG_FILE, index=False, encoding='utf-8-sig')

print(f'\n状态已保存：{os.path.basename(STATE_FILE)}；触发记录追加至 {os.path.basename(LOG_FILE)}')

# 供定时任务/通知读取的摘要
with open(os.path.join(base_dir, '监控摘要.txt'), 'w', encoding='utf-8') as f:
    f.write(f'asof={last_date.date()}\nactions={len(actions)}\n')
    for c, a, dd in actions:
        f.write(f'{a}|{c}|{dd:.2%}\n')

# ---------- 6. 推送卡片（表格：净值 / 份额 / 金额 / 调仓） ----------
try:
    import importlib
    push_table_card = importlib.import_module('推送').push_table_card

    table_rows, total_amt = [], 0.0
    total_day_pnl, total_month_pnl = 0.0, 0.0   # 用于合计行的日收益/本月收益
    sell_amt = 0.0                     # 权益类卖出所得（决定债券行的买入份额）
    for _, r in res.iterrows():
        c = r['基金']
        nav = nav_at(c, last_date)
        sh = funds[c].get('shares', np.nan)
        amt = (sh * nav) if (pd.notna(sh) and pd.notna(nav)) else np.nan
        act = [a for (cc, a, _) in actions if cc == c]
        dq_sh, dq_amt = 0.0, 0.0
        if act and act[0] == '清仓' and pd.notna(amt):
            dq_sh, dq_amt = -sh, -amt
        elif act and act[0] == '减半仓' and pd.notna(amt):
            dq_sh, dq_amt = -sh / 2, -amt / 2
        elif act and act[0] == '回补' and pd.notna(nav) and CAPITAL > 0:
            tgt = W[c] * CAPITAL
            dq_sh = tgt / nav - (sh if pd.notna(sh) else 0)
            dq_amt = tgt - (amt if pd.notna(amt) else 0)
        if act and act[0] in ('清仓', '减半仓') and c != BOND and pd.notna(amt):
            sell_amt += -dq_amt
        if pd.notna(amt):
            total_amt += amt
        # 日涨跌 / 日收益 / 本月涨跌 / 本月收益（按文档第十八节口径）
        nav_yest = nav_prev(c, last_date)
        nav_ms = nav_month_start(c, last_date)
        day_chg = (nav / nav_yest - 1) if (pd.notna(nav) and pd.notna(nav_yest) and nav_yest > 0) else np.nan
        day_pnl = sh * (nav - nav_yest) if (pd.notna(sh) and pd.notna(nav) and pd.notna(nav_yest)) else np.nan
        month_chg = (nav / nav_ms - 1) if (pd.notna(nav) and pd.notna(nav_ms) and nav_ms > 0) else np.nan
        month_pnl = sh * (nav - nav_ms) if (pd.notna(sh) and pd.notna(nav) and pd.notna(nav_ms)) else np.nan
        if pd.notna(day_pnl):
            total_day_pnl += day_pnl
        if pd.notna(month_pnl):
            total_month_pnl += month_pnl
        table_rows.append({
            'fund': c,
            'nav': f'{nav:.4f}' if pd.notna(nav) else '—',
            'day_chg': f'{day_chg:+.2%}' if pd.notna(day_chg) else '—',
            'day_pnl': f'{day_pnl:+,.2f}' if pd.notna(day_pnl) else '—',
            'month_chg': f'{month_chg:+.2%}' if pd.notna(month_chg) else '—',
            'month_pnl': f'{month_pnl:+,.2f}' if pd.notna(month_pnl) else '—',
            'shares': f'{sh:.2f}' if pd.notna(sh) else '—',
            'amount': f'{amt:,.2f}' if pd.notna(amt) else '—',
            'dshares': f'{dq_sh:+,.2f}' if abs(dq_sh) > 1e-9 else '—',
            'damount': f'{dq_amt:+,.2f}' if abs(dq_amt) > 1e-9 else '—',
            'dd': r['自高点回撤'],
            'peak_date': peak_date(c),
        })

    # 止损资金转入债券：在债券行体现买入份额；债券自身已清仓则转现金（不买债券）
    if sell_amt > 0 and bond_on:
        for x in table_rows:
            if x['fund'] == BOND:
                bn = nav_at(BOND, last_date)
                if pd.notna(bn) and bn > 0:
                    x['dshares'] = f'{sell_amt / bn:+,.2f}'
                    x['damount'] = f'{sell_amt:+,.2f}'

    # 份额台账随动作更新：清仓清零、减半卖出份额的一半、回补恢复目标份额、债券接收止损资金
    for c, a, _ in actions:
        nav = nav_at(c, last_date)
        sh = funds[c].get('shares', 0.0)
        if a == '清仓':
            funds[c]['shares'] = 0.0
        elif a == '减半仓':
            funds[c]['shares'] = sh / 2
        elif a == '回补' and CAPITAL > 0:
            if pd.notna(nav) and nav > 0:
                funds[c]['shares'] = W[c] * CAPITAL / nav
    if sell_amt > 0 and bond_on:
        bn = nav_at(BOND, last_date)
        if pd.notna(bn) and bn > 0:
            funds[BOND]['shares'] = funds[BOND].get('shares', 0.0) + sell_amt / bn
    if actions:
        with open(STATE_FILE, 'w', encoding='utf-8') as f:
            json.dump(state, f, ensure_ascii=False, indent=2)

    # 合计行：持仓金额合计 + 调仓金额净额 + 日/月收益与涨跌幅（按收益 ÷ 期初市值）
    net_dq = 0.0
    for x in table_rows:
        if x['damount'] != '—':
            net_dq += float(x['damount'].replace(',', ''))
    prev_total_amt = total_amt - total_day_pnl          # 昨日总市值
    month_start_total_amt = total_amt - total_month_pnl # 上月末总市值
    total_day_chg = total_day_pnl / prev_total_amt if prev_total_amt > 0 else np.nan
    total_month_chg = total_month_pnl / month_start_total_amt if month_start_total_amt > 0 else np.nan
    table_rows.append({
        'fund': '合计', 'nav': '—',
        'day_chg': f'{total_day_chg:+.2%}' if pd.notna(total_day_chg) else '—',
        'day_pnl': f'{total_day_pnl:+,.2f}',
        'month_chg': f'{total_month_chg:+.2%}' if pd.notna(total_month_chg) else '—',
        'month_pnl': f'{total_month_pnl:+,.2f}',
        'shares': '—',
        'amount': f'{total_amt:,.2f}',
        'dshares': '—',
        'damount': f'{net_dq:+,.2f}' if abs(net_dq) > 1e-9 else '—',
        'dd': '—',
        'peak_date': '—',
    })

    if actions:
        intro = f'**触发 {len(actions)} 项，需当日处理**'
    else:
        intro = '**未触发止损线**（权益 10% 减半 / 15% 清仓，债券 3% 清仓），全部持仓继续持有。'
    if actions and not bond_on:
        intro += '\n债券自身已清仓，止损资金转现金（收益按 0 计），不再买入债券。'
    if CAPITAL > 0:
        intro += f'\n数据日 {last_date.date()}　总资产 {CAPITAL:,.2f} 元　持仓市值 {total_amt:,.2f} 元'
    else:
        intro += f'\n数据日 {last_date.date()}　（未配置总资产，金额列留空）'

    footer = (f'高点自 {funds[W.index[0]].get("since", "—")} 起连续跟踪，调仓不清零；'
              '权益回撤 10% 减半 / 15% 清仓（资金转债券），债券 3% 清仓转现金，'
              '冷静 10 个交易日、回补需近 20 日收益 >0 且回撤修复至 -5% 以内。'
              '调仓份额为正表示买入，为负表示卖出。')

    push_table_card(
        f'组合止损监控 · {last_date.date()}',
        intro=intro,
        columns=[('fund', '基金', 'auto'), ('nav', '净值', 'auto'),
                 ('day_chg', '日涨跌', 'auto'), ('day_pnl', '日收益', 'auto'),
                 ('month_chg', '本月涨跌', 'auto'), ('month_pnl', '本月收益', 'auto'),
                 ('shares', '份额', 'auto'), ('amount', '持仓金额', 'auto'),
                 ('dshares', '调仓份额', 'auto'), ('damount', '调仓金额', 'auto'),
                 ('dd', '回撤', 'auto'), ('peak_date', '最高点日期', 'auto')],
        rows=table_rows,
        footer=footer,
        template='red' if actions else 'green',
    )
except Exception as e:
    print(f'[推送] 跳过：{e}')
