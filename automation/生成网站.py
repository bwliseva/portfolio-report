# -*- coding: utf-8 -*-
"""生成可公网访问的组合投资网站

输入：
  每日基金净值_自动_*.xlsx    （区间回报 / 单位净值）
  投资组合分析明细.xlsx        （最优配置权重 / 风险平价配置权重）
  持仓配置.json / 止损监控状态.json
输出：
  site/index.html            实时看板（曲线 / 指数多选 / 持仓与止损状态）
  site/report.html           最新一期月度报告（单文件版）
  site/assets/data.js        站点数据
  site/assets/echarts.min.js

说明：
  1) 曲线起点为当年第一个交易日（今年以来，起点 0%）
  2) 「MV」「RP」按 24 只基金模型池，每月首个交易日用当时可得数据重算权重并调回；
     「真实持仓」按基准日实际份额固定权重
  3) 止损曲线按「权益类自高点回撤 ≥10% 减至 70%、≥18% 减至 40%、≥25% 清仓；债券 ≥3% 清仓转现金；
     冷静 20 个交易日、近 20 日收益 > 0 且回撤修复至 -5% 以内回补」逐日模拟
"""
import os
import sys
import glob
import json
import shutil
import datetime

import numpy as np
import pandas as pd
from scipy.optimize import minimize

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

PLAN = '投资组合分析明细.xlsx'
CFG = '持仓配置.json'
STATE = '止损监控状态.json'
IDX_CACHE = '指数行情_网站.csv'
OUT = os.path.join(BASE, 'site')
ASSETS = os.path.join(OUT, 'assets')

BOND = '华泰保兴安悦债券C'
BONDS = (BOND, '国泰中债1-3年国开债A')   # 止损资金平均分到这两只债券
FIXED_WEIGHTS = {BOND: 0.20, '国泰中债1-3年国开债A': 0.20}
# 权益阶梯减仓：(自高点回撤阈值, 目标仓位)
LEVELS = [(0.10, 0.70), (0.18, 0.40), (0.25, 0.00)]
BOND_DD = 0.03                                 # 债券清仓
RECOVER_DD, RECOVER_WIN, COOLDOWN = -0.05, 20, 20
MIN_HISTORY = 260                              # 面板对齐：至少 260 个日度样本才保留该基金

INDEXES = [('上证指数', '1.000001'), ('深证成指', '0.399001'), ('沪深300', '1.000300'),
           ('中证500', '1.000905'), ('中证1000', '1.000852'), ('创业板指', '0.399006'),
           ('科创50', '1.000688'), ('中证红利', '1.000922'), ('恒生指数', '100.HSI')]
DEFAULT_INDEXES = ['沪深300', '中证1000']


# ---------------------------------------------------------------- 数据读取
def read_weights(sheet):
    df = pd.read_excel(PLAN, sheet_name=sheet)
    df = df.rename(columns={df.columns[0]: '基金', df.columns[1]: '权重'})
    df = df.dropna(subset=['基金'])
    out = {}
    for _, r in df.iterrows():
        try:
            v = float(r['权重'])
        except (TypeError, ValueError):
            continue
        if abs(v) > 1e-6:
            out[str(r['基金'])] = v
    total = sum(out.values())
    if total > 0:
        out = {k: v / total for k, v in out.items()}
    return out


def load_nav():
    files = sorted(glob.glob(os.path.join(BASE, '每日基金净值_自动_*.xlsx')))
    path = files[-1] if files else os.path.join(BASE, '每日基金净值.xls')
    raw = pd.read_excel(path, sheet_name='区间回报', index_col=0)
    raw = raw[~raw.index.astype(str).str.contains('数据来源')]
    raw.index = pd.to_datetime(raw.index, errors='coerce')
    raw = raw[raw.index.notna()].sort_index().apply(pd.to_numeric, errors='coerce').ffill()
    raw = raw.dropna(how='all')
    # 面板对齐：各基金按自身成立日进入，不做整表 dropna（否则会被最新成立的基金截断历史）
    raw = raw.loc[:, raw.notna().sum() >= MIN_HISTORY]
    prices = 1 + raw / 100.0
    navs = None
    try:
        navs = pd.read_excel(path, sheet_name='单位净值', index_col=0)
        navs.index = pd.to_datetime(navs.index, errors='coerce')
        navs = navs[navs.index.notna()].apply(pd.to_numeric, errors='coerce').sort_index().ffill()
    except Exception as e:
        print(f'[警告] 单位净值读取失败：{e}')
    return prices, navs, os.path.basename(path)


def nav_at(navs, c, d):
    if navs is None or c not in navs.columns:
        return np.nan
    s = navs[c].dropna()
    s = s[s.index <= pd.Timestamp(d)]
    return float(s.iloc[-1]) if len(s) else np.nan


# ---------------------------------------------------------------- 月度重优化
_TARGET_CACHE = {}


def _opt_mv(mu, cov, max_weight=0.3, n_starts=60, fixed_weights=None):
    """最大夏普：凸化精确求解；fixed_weights 为 {列索引: 固定目标权重}。"""
    mu = np.asarray(mu, dtype=float)
    cov = np.asarray(cov, dtype=float)
    n = len(mu)
    ex = mu - 0.02
    fixed_weights = fixed_weights or {}
    if np.max(ex) <= 0:
        cons = [{'type': 'eq', 'fun': lambda w: np.sum(w) - 1}]
        cons += [{'type': 'eq', 'fun': lambda w, i=i, t=t: w[i] - t}
                 for i, t in fixed_weights.items()]
        res = minimize(lambda w: w @ cov @ w, np.full(n, 1 / n), method='SLSQP',
                       bounds=[(0, max_weight)] * n, constraints=cons,
                       options={'ftol': 1e-14, 'maxiter': 2000, 'disp': False})
        return res.x if res.success else None
    y0 = np.clip(ex, 0, None)
    y0 = y0 / (ex @ y0)
    cons = [{'type': 'eq', 'fun': lambda y: ex @ y - 1},
            {'type': 'ineq', 'fun': lambda y: max_weight * np.sum(y) - y}]
    cons += [{'type': 'eq', 'fun': lambda y, i=i, t=t: y[i] - t * np.sum(y)}
             for i, t in fixed_weights.items()]
    res = minimize(lambda y: y @ cov @ y, y0, method='SLSQP',
                   bounds=[(0, None)] * n, constraints=cons,
                   options={'ftol': 1e-14, 'maxiter': 2000, 'disp': False})
    if not res.success or np.sum(res.x) <= 0:
        return None
    return res.x / np.sum(res.x)


def _opt_rp(cov):
    """风险平价（ERC）：坐标下降精确求解，无随机起点，结果稳定。"""
    S = np.asarray(cov, dtype=float)
    n = S.shape[0]
    d = np.diag(S).copy()
    w = 1.0 / np.sqrt(d)
    for _ in range(5000):
        wn = w.copy()
        for i in range(n):
            # Gauss-Seidel 顺序更新，避免高条件数下 Jacobi 迭代溢出
            b = float(S[i] @ wn - S[i, i] * wn[i])
            disc = b * b + 4 * d[i] / n
            wn[i] = (-b + np.sqrt(disc)) / (2 * d[i]) if disc > 0 else wn[i]
        if not np.all(np.isfinite(wn)):
            return None
        if np.max(np.abs(wn - w)) < 1e-15:
            w = wn
            break
        w = wn
    return w / w.sum()


def _park_bond(amt, row, shares, pos):
    """把金额平均分到仍在仓的债券；无可用债券时返回未分配金额（转现金）。"""
    if amt <= 0:
        return 0.0
    avail = [b for b in BONDS if b in shares and pos.get(b, 0.0) > 0 and row[b] > 0]
    if not avail:
        return amt
    each = amt / len(avail)
    for b in avail:
        shares[b] += each / row[b]
    return 0.0


def _draw_bonds(need, row, shares, pos):
    """从仍在仓的债券平均赎回，返回实际到手金额。"""
    avail = [b for b in BONDS if b in shares and pos.get(b, 0.0) > 0
             and shares[b] > 0 and row[b] > 0]
    if not avail:
        return 0.0
    got, per = 0.0, need / len(avail)
    for b in avail:
        amt = min(per, shares[b] * row[b])
        shares[b] -= amt / row[b]
        got += amt
    for b in avail:                          # 平均取用后仍有缺口时按剩余可用额补齐
        if got >= need - 1e-9:
            break
        amt = min(need - got, shares[b] * row[b])
        if amt > 0:
            shares[b] -= amt / row[b]
            got += amt
    return got


def monthly_target(prices, dt, mode):
    """用 dt 之前（不含当日）的数据重算目标权重。

    期望收益：近 3 年月末收益均值 × 12；协方差：近 1 年日收益 × 252。
    数据不足或优化失败返回 None（调用方沿用上期权重）。
    """
    key = (mode, pd.Timestamp(dt))
    if key in _TARGET_CACHE:
        return _TARGET_CACHE[key]
    # 004253 仅供实际持仓止损监控，不进入 MV/RP 月度优化池
    model_prices = prices.drop(columns=['国泰黄金ETF联接C'], errors='ignore')
    hist = model_prices.loc[model_prices.index < dt].ffill()
    if len(hist) < 260:
        return None
    dret = hist.pct_change(fill_method=None).dropna(how='all')
    dret = dret.loc[dret.index >= hist.index[-1] - pd.DateOffset(years=1)]
    dret = dret.dropna(axis=1)          # 只保留估计窗口内数据完整的基金（面板对齐）
    dret = dret.loc[:, dret.std() > 1e-8]
    if len(dret) < 120 or dret.shape[1] < 2:
        return None
    if mode == 'mv' and not all(c in dret.columns for c in FIXED_WEIGHTS):
        return None
    cov = dret.cov().values * 252 + np.eye(dret.shape[1]) * 1e-6
    if mode == 'mv':
        mp = hist.loc[hist.index >= hist.index[-1] - pd.DateOffset(years=3)]
        try:
            mp = mp.resample('ME').last()
        except ValueError:
            mp = mp.resample('M').last()
        mr = mp.pct_change(fill_method=None).dropna(how='all')
        mu = (mr.mean() * 12).reindex(dret.columns).fillna(0.0).values
        fixed_idx = {dret.columns.get_loc(c): v for c, v in FIXED_WEIGHTS.items()}
        w = _opt_mv(mu, cov, max_weight=0.3, fixed_weights=fixed_idx)
        if w is None:
            w = _opt_mv(mu, cov, max_weight=1.0, fixed_weights=fixed_idx)
    else:
        w = _opt_rp(cov)
    if w is None:
        return None
    s = pd.Series(w, index=dret.columns).clip(lower=0)
    s = s[s > 1e-6]
    if s.sum() <= 0:
        return None
    s = s / s.sum()
    _TARGET_CACHE[key] = s
    return s


# ---------------------------------------------------------------- 曲线计算
def curve_buyhold(prices, w, start):
    """按目标权重建仓后固定份额持有（不再平衡，权重随市值漂移）"""
    cols = [c for c in w if c in prices.columns]
    if not cols:
        return None
    sub = prices.loc[prices.index >= start, cols].ffill()
    sub = sub[sub.notna().all(axis=1)]
    if len(sub) < 2:
        return None
    ww = pd.Series({c: w[c] for c in cols}, dtype=float)
    ww = ww / ww.sum()
    lvl = (sub * ww).sum(axis=1)
    return lvl / lvl.iloc[0] - 1


def curve_monthly(prices, w, start, stop=False, bond=BOND, mode=None):
    """按目标权重建仓，每月首个交易日调回目标权重；stop=True 时叠加逐日止损。

    mode='mv'/'rp' 时，每月用当时可得数据重新跑优化得到新目标权重；
    mode=None 时沿用传入的固定权重。
    止损规则：权益 ≥10% 减至 70%、≥18% 减至 40%、≥25% 清仓；债券 ≥3% 清仓转现金。
    权益止损腾出的资金平均转入两只债券；两只都已清仓时转现金（收益按 0 计）。
    """
    cols = [c for c in prices.columns] if mode else [c for c in w if c in prices.columns]
    if not cols or (stop and bond not in prices.columns):
        return None
    sub = prices.loc[prices.index >= start, cols].ffill().bfill()
    if len(sub) < 3:
        return None
    if mode:
        t0 = monthly_target(prices, sub.index[0], mode)
        ww = t0.reindex(cols).fillna(0.0) if t0 is not None else pd.Series({c: w.get(c, 0.0) for c in cols})
    else:
        ww = pd.Series({c: w.get(c, 0.0) for c in cols}, dtype=float)
    ww = ww / ww.sum()
    target = ww.copy()
    p0 = sub.iloc[0]

    shares = {c: ww[c] / p0[c] for c in cols}
    pos = {c: 1.0 for c in cols}
    peak = {c: p0[c] for c in cols}
    last_act = {c: None for c in cols}
    cash = 0.0
    nav = []
    prev_month = None

    for i in range(len(sub)):
        row = sub.iloc[i]
        month = sub.index[i].to_period('M')

        # 0) 每月首个交易日：重算目标权重并调回（止损中的基金按剩余仓位比例配置）
        if prev_month is not None and month != prev_month:
            if mode:
                new_t = monthly_target(prices, sub.index[i], mode)
                if new_t is not None:
                    new_t = new_t.reindex(cols).fillna(0.0)
                    if new_t.sum() > 0:
                        new_t = new_t / new_t.sum()
                        for c in cols:
                            if target[c] <= 1e-9 and new_t[c] > 1e-9:
                                pos[c] = 1.0
                                peak[c] = row[c]
                                last_act[c] = None
                            elif target[c] > 1e-9 and new_t[c] <= 1e-9:
                                pos[c] = 1.0
                                last_act[c] = None
                        target = new_t
            total = cash + sum(shares[c] * row[c] for c in cols)
            investable = max(total - cash, 0.0)   # 原有现金不参与重估
            for c in cols:
                share = pos[c] if stop else 1.0
                shares[c] = target[c] * share * investable / row[c] if row[c] > 0 else 0.0
            if stop:
                lost = investable - sum(shares[c] * row[c] for c in cols)
                # 债券自身止损的部分留作现金，权益止损部分才平均补进仍在仓的债券
                bond_lost = investable * sum(target[b] * (1 - pos[b]) for b in BONDS if b in cols)
                eq_lost = max(lost - bond_lost, 0.0)
                if eq_lost > 1e-12:
                    _park_bond(eq_lost, row, shares, pos)
                cash = max(total - sum(shares[c] * row[c] for c in cols), 0.0)
        prev_month = month

        if stop:
            # 1) 更新高点
            for c in cols:
                peak[c] = max(peak[c], row[c])

            # 2) 两只债券均先判：回撤 ≥3% 清仓，资金转现金
            for b in BONDS:
                if pos[b] > 0 and row[b] / peak[b] - 1 <= -BOND_DD:
                    cash += shares[b] * row[b]
                    shares[b] = 0.0
                    pos[b] = 0.0
                    last_act[b] = i

            # 3) 权益类：阶梯减仓（≥10%→70%、≥18%→40%、≥25%→清仓），腾出的资金进债券
            for c in cols:
                if c in BONDS:
                    continue
                dd = row[c] / peak[c] - 1
                want = 1.0
                for thr, frac in LEVELS:
                    if dd <= -thr + 1e-12:
                        want = min(want, frac)
                if pos[c] <= want + 1e-9:
                    continue
                amt = shares[c] * row[c] * (pos[c] - want) / pos[c]
                shares[c] *= want / pos[c]
                pos[c] = want
                if _park_bond(amt, row, shares, pos) > 0:
                    cash += amt
                last_act[c] = i

            total = cash + sum(shares[c] * row[c] for c in cols)

            # 4) 回补：冷静期满 且 近 20 日收益 > 0 且 回撤修复至 -5% 以内 → 恢复满仓
            for c in cols:
                if pos[c] >= 1.0 or last_act[c] is None or (i - last_act[c]) < COOLDOWN:
                    continue
                dd = row[c] / peak[c] - 1
                r20 = row[c] / sub.iloc[i - RECOVER_WIN][c] - 1 if i >= RECOVER_WIN else np.nan
                if np.isnan(r20) or r20 <= 0 or dd <= RECOVER_DD:
                    continue
                need = target[c] * total - shares[c] * row[c]
                if need <= 0:
                    pos[c] = 1.0
                    last_act[c] = None
                    continue
                if c in BONDS:                       # 债券只能用现金买回（此前资金已转现金）
                    got = min(need, cash)
                    cash -= got
                    shares[c] += got / row[c]
                else:
                    got = _draw_bonds(need, row, shares, pos)
                    if got < need:
                        amt = min(need - got, cash)
                        cash -= amt
                        got += amt
                    shares[c] += got / row[c]
                pos[c] = 1.0
                last_act[c] = None
                total = cash + sum(shares[x] * row[x] for x in cols)

        nav.append(cash + sum(shares[c] * row[c] for c in cols))

    s = pd.Series(nav, index=sub.index, dtype=float)
    return s / s.iloc[0] - 1


def curve_stop(prices, w, start, bond=BOND):
    """逐日模拟新规则止损：权益 ≥10% 减至 70%、≥18% 减至 40%、≥25% 清仓；债券 ≥3% 清仓转现金。

    权益止损腾出的资金平均转入两只债券；两只都已清仓时转现金（收益按 0 计）。
    """
    cols = [c for c in w if c in prices.columns]
    if not cols or bond not in prices.columns:
        return None
    sub = prices.loc[prices.index >= start, cols].ffill().bfill()
    if len(sub) < 3:
        return None
    ww = pd.Series({c: w[c] for c in cols}, dtype=float)
    ww = ww / ww.sum()
    p0 = sub.iloc[0]

    shares = {c: ww[c] / p0[c] for c in cols}
    pos = {c: 1.0 for c in cols}
    peak = {c: p0[c] for c in cols}
    last_act = {c: None for c in cols}
    cash = 0.0
    nav = []

    for i in range(len(sub)):
        row = sub.iloc[i]
        # 1) 更新高点
        for c in cols:
            peak[c] = max(peak[c], row[c])

        # 2) 两只债券均先判：回撤 ≥3% 清仓，资金转现金
        for b in BONDS:
            if b in cols and pos[b] > 0 and row[b] / peak[b] - 1 <= -BOND_DD:
                cash += shares[b] * row[b]
                shares[b] = 0.0
                pos[b] = 0.0
                last_act[b] = i

        # 3) 权益类：阶梯减仓（≥10%→70%、≥18%→40%、≥25%→清仓），腾出的资金进债券
        for c in cols:
            if c in BONDS:
                continue
            dd = row[c] / peak[c] - 1
            want = 1.0
            for thr, frac in LEVELS:
                if dd <= -thr + 1e-12:
                    want = min(want, frac)
            if pos[c] <= want + 1e-9:
                continue
            amt = shares[c] * row[c] * (pos[c] - want) / pos[c]
            shares[c] *= want / pos[c]
            pos[c] = want
            if _park_bond(amt, row, shares, pos) > 0:
                cash += amt
            last_act[c] = i

        total = cash + sum(shares[c] * row[c] for c in cols)

        # 4) 回补：冷静期满 且 近 20 日收益 > 0 且 回撤修复至 -5% 以内 → 恢复满仓
        for c in cols:
            if pos[c] >= 1.0 or last_act[c] is None or (i - last_act[c]) < COOLDOWN:
                continue
            dd = row[c] / peak[c] - 1
            r20 = row[c] / sub.iloc[i - RECOVER_WIN][c] - 1 if i >= RECOVER_WIN else np.nan
            if np.isnan(r20) or r20 <= 0 or dd <= RECOVER_DD:
                continue
            need = ww[c] * total - shares[c] * row[c]
            if need <= 0:
                pos[c] = 1.0
                last_act[c] = None
                continue
            if c in BONDS:                       # 债券只能用现金买回（此前资金已转现金）
                got = min(need, cash)
                cash -= got
                shares[c] += got / row[c]
            else:
                got = _draw_bonds(need, row, shares, pos)
                if got < need:
                    amt = min(need - got, cash)
                    cash -= amt
                    got += amt
                shares[c] += got / row[c]
            pos[c] = 1.0
            last_act[c] = None
            total = cash + sum(shares[x] * row[x] for x in cols)

        nav.append(total)

    s = pd.Series(nav, index=sub.index, dtype=float)
    return s / s.iloc[0] - 1


def fetch_index(secid, beg='20260101'):
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
    rows = {}
    for line in (r.json().get('data') or {}).get('klines') or []:
        p = line.split(',')
        if len(p) >= 2:
            try:
                rows[pd.Timestamp(p[0])] = float(p[1])
            except ValueError:
                continue
    if not rows:
        raise RuntimeError('接口返回空数据')
    return pd.Series(rows).sort_index()


def load_indexes(as_of):
    names = [n for n, _ in INDEXES]
    cached = None
    if os.path.exists(IDX_CACHE):
        try:
            c = pd.read_csv(IDX_CACHE, encoding='utf-8-sig')
            c['日期'] = pd.to_datetime(c['日期'], errors='coerce')
            c = c.dropna(subset=['日期']).set_index('日期').sort_index()
            cached = c.apply(pd.to_numeric, errors='coerce')
        except Exception as e:
            print(f'[警告] {IDX_CACHE} 读取失败：{e}')
    if (cached is not None and all(n in cached.columns for n in names)
            and cached.index.max() >= as_of):
        return cached[names], '缓存'
    data = {}
    for n, sid in INDEXES:
        try:
            data[n] = fetch_index(sid)
        except Exception as e:
            print(f'[警告] 指数「{n}」抓取失败：{e}')
            if cached is not None and n in cached.columns:
                data[n] = cached[n].dropna()
    if not data:
        return None, ''
    df = pd.DataFrame(data).sort_index()
    df.index.name = '日期'
    try:
        df.to_csv(IDX_CACHE, encoding='utf-8-sig')
    except Exception as e:
        print(f'[警告] {IDX_CACHE} 写入失败：{e}')
    return df[[n for n in names if n in df.columns]], '东方财富'


# ---------------------------------------------------------------- 主流程
prices, navs, nav_file = load_nav()
as_of = prices.index[-1]
print(f'净值数据：{nav_file}，最新 {as_of.date()}，{prices.shape[1]} 只基金')

mv = read_weights('最优配置权重')
rp = read_weights('风险平价配置权重')
mv_nz = {k: v for k, v in mv.items() if v > 1e-4}
rp_nz = {k: v for k, v in rp.items() if v > 1e-4}

with open(CFG, encoding='utf-8') as f:
    cfg = json.load(f)
with open(STATE, encoding='utf-8') as f:
    state = json.load(f)
funds_state = state.get('funds', {})

start = prices.index[prices.index >= pd.Timestamp(f'{as_of.year}-01-01')]
start = start[0] if len(start) else prices.index[0]

# 真实持仓（状态中有正份额的基金）
held = []
for c, st in funds_state.items():
    try:
        sh = float(st.get('shares', 0) or 0)
    except (TypeError, ValueError):
        sh = 0.0
    if sh > 0 and c in prices.columns:
        held.append((c, sh))
w_real = {}
for c, sh in held:
    nv = nav_at(navs, c, as_of)
    if pd.notna(nv) and nv > 0:
        w_real[c] = sh * nv
if w_real:
    tot = sum(w_real.values())
    w_real = {k: v / tot for k, v in w_real.items()}

# 曲线（MV / RP = 24 只模型池，每月重算权重并调回；止损按月调+日度止损）
curves = {}
curves['mv'] = curve_monthly(prices, mv_nz, start, stop=False, mode='mv')
curves['rp'] = curve_monthly(prices, rp_nz, start, stop=False, mode='rp')
curves['real'] = curve_buyhold(prices, w_real, start)
curves['mv_stop'] = curve_monthly(prices, mv_nz, start, stop=True, mode='mv')
curves['real_stop'] = curve_stop(prices, w_real, start)

common = None
for s in curves.values():
    if s is None:
        continue
    common = s.index if common is None else common.intersection(s.index)
common = common.sort_values()
dates = [d.strftime('%Y-%m-%d') for d in common]

series = {}
for k, s in curves.items():
    series[k] = [round(float(v) * 100, 2) for v in s.reindex(common)] if s is not None else [None] * len(common)

# 指数
idx_df, idx_src = load_indexes(as_of)
idx_asof = idx_df.index.max() if idx_df is not None else as_of
indexes = []
if idx_df is not None:
    for n in [x for x, _ in INDEXES]:
        if n not in idx_df.columns:
            continue
        s = idx_df[n].reindex(common).ffill().bfill()
        base = float(s.iloc[0]) if len(s) else np.nan
        indexes.append({'name': n, 'values': [round(float(v / base - 1) * 100, 2) for v in s]})

hs300 = series_val = None
for x in indexes:
    if x['name'] == '沪深300':
        hs300 = x['values'][-1]

kpi = {}
for k, vals in series.items():
    last = next((v for v in reversed(vals) if v is not None), None)
    kpi[k] = last
    kpi[k + '_ex'] = (round(last - hs300, 2) if (last is not None and hs300 is not None) else None)

# 持仓表（均值方差权重按“已持有基金”归一化到 100%；现金与其他持仓不参与模型）
rows = []
sum_w_mv = 0.0
for c, st in funds_state.items():
    try:
        sh = float(st.get('shares', 0) or 0)
    except (TypeError, ValueError):
        sh = 0.0
    if sh <= 0:
        continue
    nv = nav_at(navs, c, as_of)
    val = sh * nv if pd.notna(nv) else np.nan
    peak = float(st.get('peak') or 0)
    p_now = float(prices[c].dropna().iloc[-1]) if c in prices.columns else np.nan
    dd = (p_now / peak - 1) if (peak > 0 and pd.notna(p_now)) else np.nan
    sum_w_mv += mv_nz.get(c, 0)
    rows.append({'name': c, 'w_raw': mv_nz.get(c, 0), 'shares': sh, 'value_num': val,
                 'dd': f'{dd * 100:.2f}%' if pd.notna(dd) else '—', 'dd_num': dd,
                 'on': float(st.get('pos', 1.0 if st.get('on', True) else 0.0)) > 0,
                 'pos': float(st.get('pos', 1.0 if st.get('on', True) else 0.0)),
                 'kind': 'fund'})

cash = float(cfg.get('现金', 0) or 0)
if cash > 0:
    rows.append({'name': '现金', 'w_raw': 0.0, 'shares': np.nan, 'value_num': cash,
                 'dd': '—', 'dd_num': np.nan, 'on': True, 'kind': 'cash'})
for o in cfg.get('其他持仓', []):
    v = float(o.get('市值', 0) or 0)
    rows.append({'name': f"{o.get('代码', '')} {o.get('名称', '')}".strip(), 'w_raw': 0.0,
                 'shares': float(o.get('份额', 0) or 0), 'value_num': v,
                 'dd': '—', 'dd_num': np.nan, 'on': True, 'kind': 'other'})

rows.sort(key=lambda r: -(r['value_num'] if pd.notna(r['value_num']) else 0))
holdings = []
total_val = 0.0
for r in rows:
    w = r['w_raw'] / sum_w_mv if sum_w_mv > 0 else 0.0
    if pd.notna(r['value_num']):
        total_val += r['value_num']
    holdings.append({
        'name': r['name'],
        'w_mv': f'{w * 100:.2f}%',
        'value': f"{r['value_num']:,.2f}" if pd.notna(r['value_num']) else '—',
        'shares': f"{r['shares']:,.2f}" if pd.notna(r['shares']) else '—',
        'dd': r['dd'],
        'dd_num': round(float(r['dd_num']) * 100, 2) if pd.notna(r['dd_num']) else 0,
        'on': r['on'],
        'pos': r.get('pos', 1.0),
        'kind': r['kind'],
    })

hold_sum = {'w_mv': '100.00%', 'value': f'{total_val:,.2f}', 'count': len(holdings)}

hold_names = [h['name'] for h in holdings]
subtitle = (f'持仓 {len(holdings)} 项（基准日 {cfg.get("基准日", "—")}）　·　模型池 24 只　·　'
            f'指数 {len(indexes)} 个可自由勾选')

payload = {
    'asof': as_of.strftime('%Y-%m-%d'),
    'updated': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),
    'indexAsOf': idx_asof.strftime('%Y-%m-%d'),
    'subtitle': subtitle,
    'dates': dates,
    'series': series,
    'indexes': indexes,
    'defaultIndexes': DEFAULT_INDEXES,
    'kpi': kpi,
    'holdings': holdings,
    'holdSummary': hold_sum,
    'note': ('数据来源：东方财富（基金净值 / 指数 K 线）。累计收益率以区间首日为 0%，'
             '不含申购赎回费用；止损曲线为规则模拟，不代表实际成交。仅供个人研究参考。'),
}

# ---------------------------------------------------------------- 输出
os.makedirs(ASSETS, exist_ok=True)
with open(os.path.join(ASSETS, 'data.js'), 'w', encoding='utf-8') as f:
    f.write('window.SITE = ' + json.dumps(payload, ensure_ascii=False) + ';\n')

shutil.copyfile(os.path.join(BASE, '网站模板.html'), os.path.join(OUT, 'index.html'))

# 给 data.js 加版本号，避免浏览器缓存旧数据（否则页面仍显示上一次的数据日）
_index_path = os.path.join(OUT, 'index.html')
_ver = datetime.datetime.now().strftime('%Y%m%d%H%M')
with open(_index_path, encoding='utf-8') as f:
    _html = f.read()
_html = _html.replace('assets/data.js"', f'assets/data.js?v={_ver}"')
with open(_index_path, 'w', encoding='utf-8') as f:
    f.write(_html)

echarts_files = [p for p in glob.glob(os.path.join(BASE, '**', 'echarts.min.js'), recursive=True)
                 if os.path.commonpath([os.path.abspath(p), OUT]) != os.path.abspath(OUT)]
if echarts_files:
    shutil.copyfile(sorted(echarts_files, key=os.path.getsize)[-1], os.path.join(ASSETS, 'echarts.min.js'))
else:
    print('[警告] 未找到 echarts.min.js，网站图表无法显示')

reports = sorted(glob.glob(os.path.join(BASE, 'portfolio-report-*', 'portfolio-report-*-单文件.html')))
if reports:
    shutil.copyfile(reports[-1], os.path.join(OUT, 'report.html'))
    print(f'报告页：{os.path.basename(reports[-1])}')
else:
    print('[警告] 未找到单文件报告，report.html 未生成')

print(f'\n网站已生成：{OUT}')
print(f'数据日 {payload["asof"]}（指数 {payload["indexAsOf"]}）｜{len(dates)} 个交易日')
for k in ['mv_stop', 'mv', 'real_stop', 'real', 'rp']:
    ex = kpi.get(k + '_ex')
    if ex is None:
        print(f'  {k:10s} {kpi[k]:+7.2f}%')
    else:
        print(f'  {k:10s} {kpi[k]:+7.2f}%（较沪深300 {ex:+.2f}%）')
print(f'  沪深300   {hs300:+7.2f}%｜指数源 {idx_src}')

# ---------------------------------------------------------------- 自动部署
deploy_cfg = os.path.join(BASE, '网站配置.json')
if os.path.exists(deploy_cfg):
    try:
        with open(deploy_cfg, encoding='utf-8') as f:
            d = json.load(f)
        if d.get('token') and d.get('owner'):
            import subprocess
            r = subprocess.run([sys.executable, os.path.join(BASE, '部署网站.py')],
                               capture_output=True, text=True, encoding='utf-8', errors='ignore')
            print(r.stdout.strip()[-2000:] if r.stdout else '')
            if r.returncode != 0:
                print('[部署失败]', (r.stderr or '')[-800:])
        else:
            print('[部署] 未填写 GitHub token/owner，跳过（见 网站配置.json）')
    except Exception as e:
        print(f'[部署] 跳过：{e}')

# 部署失败不影响站点生成，统一以 0 退出，避免定时任务被标记为失败
sys.exit(0)
