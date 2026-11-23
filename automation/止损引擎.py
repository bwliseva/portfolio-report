# -*- coding: utf-8 -*-
"""月度重估 + 固定份额 + 日度止损 的统一回测引擎。

与网站曲线（生成网站.py 的 curve_monthly(mode='mv', stop=True)）同一口径：
  1. 起始日按当期最优权重建仓，记录份额；
  2. 每月首个交易日用「该日之前」的数据重估最大夏普权重，并调回目标权重；
  3. 两次调仓之间份额固定，权重随净值漂移，不再每日再平衡；
  4. 日度止损：权益自高点回撤 ≥10% 减半、≥15% 清仓；债券 ≥3% 清仓；
     腾出资金平均分到两只债券，均不可用时转现金（收益按 0 计）；
  5. 冷静期满 N 个交易日，且近 20 日收益 > 0、回撤修复至 -5% 以内，回补至目标权重。
"""
import numpy as np
import pandas as pd
from scipy.optimize import minimize

RF = 0.02
MAX_W = 0.30
BOND = '华泰保兴安悦债券C'  # 保留：历史调用与回调参数默认值
BONDS = (BOND, '国泰中债1-3年国开债A')  # 止损资金平均分到这两只债券
BOND_CAPS = {BOND: 0.20, '国泰中债1-3年国开债A': 0.20}  # 单只上限 20%，不强制等于
HALF_DD, FULL_DD, BOND_DD = 0.10, 0.15, 0.03
RECOVER_DD, RECOVER_WIN, COOLDOWN = -0.05, 20, 20


def opt_mv(mu, cov, max_weight=MAX_W, fixed_weights=None, caps=None):
    """最大夏普：凸化精确求解。

    fixed_weights 为 {列索引: 固定目标权重}（等式约束）；
    caps 为 {列索引: 单只上限}（不等式约束，允许低于上限）。
    """
    mu = np.asarray(mu, dtype=float)
    cov = np.asarray(cov, dtype=float)
    n = len(mu)
    ex = mu - RF
    fixed_weights = fixed_weights or {}
    caps = caps or {}
    if np.max(ex) <= 0:
        # 无正超额收益时退化为最小方差，仍保持债券上限约束
        cons = [{'type': 'eq', 'fun': lambda w: np.sum(w) - 1}]
        cons += [{'type': 'eq', 'fun': lambda w, i=i, t=t: w[i] - t}
                 for i, t in fixed_weights.items()]
        cons += [{'type': 'ineq', 'fun': lambda w, i=i, c=c: c - w[i]}
                 for i, c in caps.items()]
        res = minimize(lambda w: w @ cov @ w, np.full(n, 1 / n), method='SLSQP',
                       bounds=[(0, max_weight)] * n, constraints=cons,
                       options={'ftol': 1e-14, 'maxiter': 2000, 'disp': False})
        return res.x if res.success else None
    y0 = np.clip(ex, 0, None)
    y0 = y0 / (ex @ y0)
    cons = [{'type': 'eq', 'fun': lambda y: ex @ y - 1},
            {'type': 'ineq', 'fun': lambda y: max_weight * np.sum(y) - y}]
    # y=w/sum(y)，故固定 w_i=t 等价于 y_i=t*sum(y)
    cons += [{'type': 'eq', 'fun': lambda y, i=i, t=t: y[i] - t * np.sum(y)}
             for i, t in fixed_weights.items()]
    # 同理 w_i<=c 等价于 y_i<=c*sum(y)
    cons += [{'type': 'ineq', 'fun': lambda y, i=i, c=c: c * np.sum(y) - y[i]}
             for i, c in caps.items()]
    res = minimize(lambda y: y @ cov @ y, y0, method='SLSQP',
                   bounds=[(0, None)] * n, constraints=cons,
                   options={'ftol': 1e-14, 'maxiter': 2000, 'disp': False})
    if not res.success or np.sum(res.x) <= 0:
        return None
    return res.x / np.sum(res.x)


def monthly_target(prices, dt, cache=None):
    """用 dt 之前（不含当日）的数据重算目标权重：近 3 年月末收益均值、近 1 年日协方差。

    数据不足或优化失败返回 None（调用方沿用上期权重）。
    """
    key = pd.Timestamp(dt)
    if cache is not None and key in cache:
        return cache[key]
    out = None
    # 004253 仅供实际持仓止损监控，不进入 MV/RP 月度优化池
    model_prices = prices.drop(columns=['国泰黄金ETF联接C'], errors='ignore')
    hist = model_prices.loc[model_prices.index < key].ffill()
    if len(hist) >= 260:
        dret = hist.pct_change(fill_method=None).dropna(how='all')
        dret = dret.loc[dret.index >= hist.index[-1] - pd.DateOffset(years=1)]
        dret = dret.dropna(axis=1)          # 只保留估计窗口内数据完整的基金（面板对齐）
        dret = dret.loc[:, dret.std() > 1e-8]
        if (len(dret) >= 120 and dret.shape[1] >= 2
                and all(c in dret.columns for c in BOND_CAPS)):
            cov = dret.cov().values * 252 + np.eye(dret.shape[1]) * 1e-6
            mp = hist.loc[hist.index >= hist.index[-1] - pd.DateOffset(years=3)]
            try:
                mp = mp.resample('ME').last()
            except ValueError:
                mp = mp.resample('M').last()
            mu = (mp.pct_change(fill_method=None).dropna(how='all').mean() * 12)
            mu = mu.reindex(dret.columns).fillna(0.0).values
            bond_cap_idx = {dret.columns.get_loc(c): w for c, w in BOND_CAPS.items()}
            w = opt_mv(mu, cov, caps=bond_cap_idx)
            if w is None:
                w = opt_mv(mu, cov, max_weight=1.0, caps=bond_cap_idx)
            if w is not None:
                s = pd.Series(w, index=dret.columns).clip(lower=0)
                s = s[s > 1e-6]
                if s.sum() > 0:
                    out = s / s.sum()
    if cache is not None:
        cache[key] = out
    return out


def simulate(prices, start, stop=True, partial=True, equity_dd=FULL_DD, bond_dd=BOND_DD,
             cooldown=COOLDOWN, cash_sink=False, ma200=False, bond=BOND, cache=None,
             levels=None):
    """份额制回测：每月重估权重并调仓，日度检查止损。

    stop=False 即「每月重估 + 固定份额」的无止损基准；
    partial=True 为 10% 减半 / equity_dd 清仓，False 为仅 equity_dd 清仓；
    levels=[(回撤阈值, 保留仓位), ...] 时改用阶梯减仓（如 [(0.10,0.6),(0.15,0.3),(0.20,0.0)]）；
    cash_sink=True 时止损资金转现金（不买债券）；ma200=True 改用 200 日均线规则。
    返回 dict(nav, half, bond_w, cash_w, events, weights, turn)。
    """
    cols = list(prices.columns)
    if any(b not in cols for b in BONDS) or len(cols) < 2:
        return None
    sub = prices.loc[prices.index >= start, cols].ffill().bfill()
    if len(sub) < 3:
        return None
    if cache is None:
        cache = {}
    t0 = monthly_target(prices, sub.index[0], cache)
    if t0 is None:
        return None
    ww = t0.reindex(cols).fillna(0.0)
    if ww.sum() <= 0:
        return None
    ww = ww / ww.sum()
    target = ww.copy()
    p0 = sub.iloc[0]
    shares = {c: (ww[c] / p0[c] if p0[c] > 0 else 0.0) for c in cols}
    pos = {c: 1.0 for c in cols}
    peak = {c: p0[c] for c in cols}
    last_act = {c: None for c in cols}
    cash = 0.0
    nav, half, bw, cw = [], [], [], []
    events = []
    weights = {sub.index[0]: ww.copy()}
    turn = {'sell': 0.0, 'buy': 0.0}
    prev_month = None
    ma = prices.rolling(200).mean() if ma200 else None

    def total_val(row):
        return cash + sum(shares[c] * row[c] for c in cols)

    def park(amt, row):
        """止损资金平均分到仍在仓的债券；两只都不可用时转现金。"""
        nonlocal cash
        if amt <= 0:
            return
        turn['sell'] += amt
        avail = [b for b in BONDS if b in cols and not cash_sink
                 and pos.get(b, 0.0) > 0 and row[b] > 0]
        if not avail:
            cash += amt
            return
        each = amt / len(avail)
        for b in avail:
            shares[b] += each / row[b]

    def take_bonds(need, row):
        """从仍在仓的债券平均赎回，返回实际到手金额。"""
        avail = [b for b in BONDS if b in cols and pos.get(b, 0.0) > 0
                 and shares[b] > 0 and row[b] > 0]
        if not avail:
            return 0.0
        got, per = 0.0, need / len(avail)
        for b in avail:
            amt = min(per, shares[b] * row[b])
            shares[b] -= amt / row[b]
            got += amt
        for b in avail:                      # 平均取用后仍有缺口时按剩余可用额补齐
            if got >= need - 1e-9:
                break
            amt = min(need - got, shares[b] * row[b])
            if amt > 0:
                shares[b] -= amt / row[b]
                got += amt
        return got

    def buy_back(c, need, row):
        """用债券（优先）与现金把 c 补到目标金额，返回实际买入金额。"""
        nonlocal cash
        if need <= 0:
            return 0.0
        got = 0.0
        turn['buy'] += 0.0
        if c not in BONDS and not cash_sink:
            got = take_bonds(need, row)
        if got < need:
            amt = min(need - got, cash)
            cash -= amt
            got += amt
        if row[c] > 0:
            shares[c] += got / row[c]
        turn['buy'] += got
        return got

    for i, dt in enumerate(sub.index):
        row = sub.iloc[i]
        month = dt.to_period('M')

        # 每月首个交易日：重估权重并调回（止损中的基金按剩余仓位比例配置）
        if prev_month is not None and month != prev_month:
            new_t = monthly_target(prices, dt, cache)
            if new_t is not None:
                new_t = new_t.reindex(cols).fillna(0.0)
                if new_t.sum() > 0:
                    new_t = new_t / new_t.sum()
                    for c in cols:
                        if target[c] <= 1e-9 and new_t[c] > 1e-9:
                            pos[c], peak[c], last_act[c] = 1.0, row[c], None
                        elif target[c] > 1e-9 and new_t[c] <= 1e-9:
                            pos[c], last_act[c] = 1.0, None
                    target = new_t
                    weights[dt] = target.copy()
            total = total_val(row)
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
                    avail = [b for b in BONDS if pos[b] > 0 and row[b] > 0]
                    if not cash_sink and avail:
                        each = eq_lost / len(avail)
                        for b in avail:
                            shares[b] += each / row[b]
                cash = max(total - sum(shares[c] * row[c] for c in cols), 0.0)
        prev_month = month

        if stop:
            for c in cols:
                peak[c] = max(peak[c], row[c])

            if ma200:
                for c in cols:
                    m = ma.loc[dt, c]
                    if not np.isfinite(m):
                        continue
                    dd = row[c] / peak[c] - 1
                    if pos[c] > 0 and row[c] < m:
                        amt = shares[c] * row[c]
                        shares[c], pos[c], last_act[c] = 0.0, 0.0, i
                        park(amt, row)
                        events.append({'日期': dt, '基金': c, '事件': '跌破MA200',
                                       '自高点回撤': dd})
                    elif pos[c] == 0 and row[c] > m:
                        need = target[c] * total_val(row) - shares[c] * row[c]
                        buy_back(c, need, row)
                        pos[c], last_act[c] = 1.0, None
                        events.append({'日期': dt, '基金': c, '事件': '站上MA200',
                                       '自高点回撤': dd})
            else:
                # 两只债券均先判：回撤 ≥bond_dd 清仓，资金转现金
                for b in BONDS:
                    if pos[b] > 0 and row[b] / peak[b] - 1 <= -bond_dd:
                        amt = shares[b] * row[b]
                        shares[b], pos[b], last_act[b] = 0.0, 0.0, i
                        cash += amt
                        turn['sell'] += amt
                        events.append({'日期': dt, '基金': b, '事件': '清仓',
                                       '自高点回撤': row[b] / peak[b] - 1})

                # 权益类：≥equity_dd 清仓优先，其次 ≥10% 减半；levels 模式下按阶梯减仓
                for c in cols:
                    if c in BONDS:
                        continue
                    dd = row[c] / peak[c] - 1
                    amt = 0.0
                    if levels:
                        want = 1.0
                        for thr, frac in levels:
                            if dd <= -thr:
                                want = min(want, frac)
                        if pos[c] > 0 and want < pos[c] - 1e-9:
                            amt = shares[c] * row[c] * (pos[c] - want) / pos[c]
                            shares[c] *= want / pos[c]
                            pos[c], last_act[c] = want, i
                            events.append({'日期': dt, '基金': c,
                                           '事件': f'减仓→{want:.0%}', '自高点回撤': dd})
                        else:
                            continue
                    elif pos[c] > 0 and dd <= -equity_dd:
                        amt = shares[c] * row[c]
                        shares[c], pos[c], last_act[c] = 0.0, 0.0, i
                        events.append({'日期': dt, '基金': c, '事件': '清仓',
                                       '自高点回撤': dd})
                    elif partial and pos[c] == 1.0 and dd <= -HALF_DD:
                        amt = shares[c] * row[c] / 2
                        shares[c] -= shares[c] / 2
                        pos[c], last_act[c] = 0.5, i
                        events.append({'日期': dt, '基金': c, '事件': '减半仓',
                                       '自高点回撤': dd})
                    else:
                        continue
                    park(amt, row)

                # 回补：冷静期满 且 近 20 日收益 > 0 且 回撤修复至 -5% 以内
                for c in cols:
                    if pos[c] >= 1.0 or last_act[c] is None:
                        continue
                    if (i - last_act[c]) < cooldown:
                        continue
                    dd = row[c] / peak[c] - 1
                    if i < RECOVER_WIN:
                        continue
                    r20 = row[c] / sub.iloc[i - RECOVER_WIN][c] - 1
                    if not (r20 > 0) or dd <= RECOVER_DD:
                        continue
                    total = total_val(row)
                    need = target[c] * total - shares[c] * row[c]
                    if c in BONDS:
                        got = min(need, cash)
                        cash -= got
                        shares[c] += got / row[c] if row[c] > 0 else 0.0
                    else:
                        buy_back(c, need, row)
                    pos[c], last_act[c] = 1.0, None
                    events.append({'日期': dt, '基金': c, '事件': '回补',
                                   '自高点回撤': dd})

        total = total_val(row)
        nav.append(total)
        half.append(sum(1 for c in cols if 0 < pos[c] < 1))
        bw.append(sum(shares[b] * row[b] for b in BONDS) / total if total > 0 else 0.0)
        cw.append(cash / total if total > 0 else 0.0)

    s = pd.Series(nav, index=sub.index, dtype=float)
    return {'nav': s / s.iloc[0], 'half': pd.Series(half, index=sub.index),
            'bond_w': pd.Series(bw, index=sub.index),
            'cash_w': pd.Series(cw, index=sub.index),
            'events': pd.DataFrame(events),
            'weights': pd.DataFrame(weights).T.fillna(0.0),
            'turn': turn}
