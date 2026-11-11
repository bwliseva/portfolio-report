# -*- coding: utf-8 -*-
"""月度调仓提醒（每月 1 日 9:00 由 Windows 计划任务触发）

逻辑：
  1) 非每月 1 日直接退出；本月已提醒过则退出
  2) 确认最新数据已覆盖上月末 → 抓数据 + 重估权重
  3) 读取新权重，与上次提醒快照对比，列出变动 ≥3pct 的基金
  4) 推送到手机，并记录本月已提醒
"""
import os, sys, json, subprocess, datetime
import pandas as pd

base_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(base_dir)

STATE_FILE = os.path.join(base_dir, '调仓提醒状态.json')
LOG_FILE = os.path.join(base_dir, '月度提醒运行日志.txt')
PLAN_FILE = os.path.join(base_dir, '投资组合分析明细.xlsx')
CHG_THRESHOLD = 0.03   # 权重变动达到 3 个百分点才提示
BOND = '华泰保兴安悦债券C'   # 止损资金默认转入的债券基金


def log(msg):
    line = f'[{datetime.datetime.now():%Y-%m-%d %H:%M}] {msg}'
    print(line)
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(line + '\n')


def main():
    today = datetime.date.today()
    # 每月 1 日固定执行（不做工作日/节假日判断；休市日看到提醒次日再操作）
    if today.day != 1 and '--force' not in sys.argv:
        return

    month_key = today.strftime('%Y-%m')
    if not _configured():
        log('未配置推送 webhook，跳过本月提醒')
        return

    state = {}
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, encoding='utf-8') as f:
            state = json.load(f)
    if state.get('last_month') == month_key and '--force' not in sys.argv:
        return

    # 数据是否已覆盖上月末
    last_auto = sorted([f for f in os.listdir(base_dir) if f.startswith('每日基金净值_自动_') and f.endswith('.xlsx')])
    if not last_auto:
        log('未找到自动数据文件，等待数据抓取后再提醒')
        return
    df = pd.read_excel(last_auto[-1], sheet_name='区间回报', index_col=0)
    df = df[~df.index.astype(str).str.contains('数据来源')]
    df.index = pd.to_datetime(df.index, errors='coerce')
    df = df[df.index.notna()].sort_index()
    data_end = df.index[-1].date()
    if data_end >= today:
        return

    log(f'开始月度调仓提醒：上月数据截止 {data_end}')

    # 抓数据 + 重估权重
    for script in ['基金数据获取.py', '更新测算.py', '生成报告.py']:
        r = subprocess.run([sys.executable, script], capture_output=True, text=True,
                           encoding='utf-8', errors='ignore')
        if r.returncode != 0:
            log(f'{script} 执行失败：{(r.stdout or "")[-200:]} {(r.stderr or "")[-300:]}')
            _push_fail(script)
            return
        log(f'{script} 完成')

    # 读取新权重
    plan = pd.read_excel(PLAN_FILE, sheet_name='最优配置权重')
    plan.columns = ['基金', '权重']
    W = plan.set_index('基金')['权重']
    W = W[W > 1e-4]
    W = W / W.sum()

    metrics = ''
    ov = pd.read_excel(PLAN_FILE, sheet_name='组合统计概览')
    try:
        m = dict(zip(ov.iloc[:, 0], ov.iloc[:, 1]))
        metrics = (f'组合预期：年化 {m.get("组合年化收益率", "")} / 波动 {m.get("组合年化波动率", "")}'
                   f' / 夏普 {m.get("夏普比率", "")}')
    except Exception:
        pass

    prev = state.get('weights') or {}
    changes = []
    for c, w in W.items():
        old = float(prev.get(c, 0))
        if abs(w - old) >= CHG_THRESHOLD:
            changes.append((c, old, w))
    changes.sort(key=lambda x: -abs(x[2] - x[1]))

    # ---------- 资金口径与份额台账 ----------
    capital = 0.0
    cfg_path = os.path.join(base_dir, '持仓配置.json')
    if os.path.exists(cfg_path):
        with open(cfg_path, encoding='utf-8') as f:
            capital = float(json.load(f).get('总资产', 0) or 0)

    nav = None
    try:
        nav = pd.read_excel(last_auto[-1], sheet_name='单位净值', index_col=0)
        nav.index = pd.to_datetime(nav.index, errors='coerce')
        nav = nav[nav.index.notna()].apply(pd.to_numeric, errors='coerce').sort_index().ffill()
    except Exception as e:
        log(f'单位净值读取失败，金额列将留空：{e}')

    def nav_at(c, d):
        if nav is None or c not in nav.columns:
            return float('nan')
        s = nav[c].dropna()
        s = s[s.index <= pd.Timestamp(d)]
        return float(s.iloc[-1]) if len(s) else float('nan')

    mstate = {}
    sf = os.path.join(base_dir, '止损监控状态.json')
    if os.path.exists(sf):
        with open(sf, encoding='utf-8') as f:
            mstate = json.load(f)
    funds = mstate.setdefault('funds', {})

    # 止损状态跨月保留：买入目标 = 总资产 × 目标权重 × pos（1.0/0.5/0.0）
    # 腾出的资金计入债券；债券自身已清仓（pos=0）则留作现金，本月不再买入
    def fund_pos(c):
        try:
            return float(funds.get(c, {}).get('pos', 1.0))
        except (TypeError, ValueError):
            return 1.0

    pos_map = {c: fund_pos(c) for c in W.index}
    bond_on = pos_map.get(BOND, 0.0) > 0 if BOND in W.index else False
    freed = capital * (1 - sum(W[c] * pos_map[c] for c in W.index)) if capital > 0 else 0.0
    freed = max(freed, 0.0)

    table_rows, new_shares, held_amt = [], {}, 0.0
    for c in W.sort_values(ascending=False).index:
        n_end = nav_at(c, data_end)
        st = funds.setdefault(c, {})
        since = st.get('since', str(data_end))
        if 'shares' not in st and capital > 0:
            # 无份额记录 = 尚未持有该基金，按真实持仓口径置 0（不假设已按目标权重建仓）
            st['shares'] = 0.0
        # 调出组合后又重新入选：止损状态清零，与回测口径一致
        if prev.get(c, 0) <= 1e-9 and W[c] > 1e-9:
            st['pos'] = 1.0
            st['on'] = True
            st['last_off'] = None
            pos_map[c] = 1.0
        pos_c = pos_map[c]
        sh_old = st.get('shares', float('nan'))
        amt_old = sh_old * n_end if pd.notna(sh_old) and pd.notna(n_end) else float('nan')
        # 止损中的基金只买回剩余仓位；债券额外承接止损腾出的资金
        eff_w = W[c] * pos_c
        if bond_on and c == BOND:
            eff_w = W[BOND] * 1.0 + (freed / capital if capital > 0 else 0.0)
        sh_new = capital * eff_w / n_end if (capital > 0 and pd.notna(n_end) and n_end > 0) else float('nan')
        dq_sh = sh_new - sh_old if pd.notna(sh_new) and pd.notna(sh_old) else float('nan')
        dq_amt = dq_sh * n_end if pd.notna(dq_sh) and pd.notna(n_end) else float('nan')
        if pd.notna(amt_old):
            held_amt += amt_old
        if pd.notna(sh_new):
            new_shares[c] = sh_new
        table_rows.append({
            'fund': c,
            'wt': f'{eff_w:.2%}',
            'nav': f'{n_end:.4f}' if pd.notna(n_end) else '—',
            'shares': f'{sh_old:,.2f}' if pd.notna(sh_old) else '—',
            'amount': f'{amt_old:,.2f}' if pd.notna(amt_old) else '—',
            'dshares': f'{dq_sh:+,.2f}' if pd.notna(dq_sh) and abs(dq_sh) > 1e-9 else '—',
            'damount': f'{dq_amt:+,.2f}' if pd.notna(dq_amt) and abs(dq_amt) > 1e-9 else '—',
        })

    # 债券已清仓时，止损资金留作现金（收益按 0 计）
    cash_amt = freed if (capital > 0 and not bond_on) else 0.0
    if cash_amt > 1e-9:
        table_rows.append({
            'fund': '现金（货币）', 'wt': f'{cash_amt / capital:.2%}', 'nav': '—',
            'shares': '—', 'amount': '—', 'dshares': '—', 'damount': '—',
        })

    # 合计行：权重合计 100% + 持仓金额合计 + 调仓金额净额
    net_dq = 0.0
    for x in table_rows:
        if x['damount'] != '—':
            net_dq += float(x['damount'].replace(',', ''))
    table_rows.append({
        'fund': '合计', 'wt': '100.00%', 'nav': '—',
        'shares': '—', 'amount': f'{held_amt:,.2f}',
        'dshares': '—',
        'damount': f'{net_dq:+,.2f}' if abs(net_dq) > 1e-9 else '—',
    })

    # 头部摘要
    if actions_txt := [f"{'加仓' if n > o else '减仓'} {c} {o:.2%}→{n:.2%}"
                       for c, o, n in changes[:4]]:
        chg_txt = '　较上期变动：' + '；'.join(actions_txt)
    elif prev:
        chg_txt = '　较上期无 3pct 以上变动'
    else:
        chg_txt = ''
    intro = (f'数据截止 {data_end}　模型已重算，请在开盘后按新权重视偏差调仓。')
    if capital > 0:
        intro += f'\n总资产 {capital:,.2f} 元　持仓市值 {held_amt:,.2f} 元{chg_txt}'
    stopped = [c for c in W.index if pos_map.get(c, 1.0) < 1]
    if stopped:
        intro += ('　止损中：' + '、'.join(f'{c}（{pos_map[c]:.0%} 仓）' for c in stopped)
                  + '，本月只按剩余仓位买入')
    if metrics:
        intro += '\n' + metrics.replace('组合预期：', '组合预期：').replace(' / ', '　')

    footer = ('权益类自高点回撤 ≥10% 减半仓、≥15% 清仓（资金转债券），债券 ≥3% 清仓转现金；'
              '冷静 10 个交易日，近 20 日收益 >0 且回撤修复至 -5% 内才回补；'
              '高点跨月连续跟踪，调仓不清零。'
              '目标权重已按止损仓位折算（半仓 0.5、清仓 0），腾出资金计入债券，债券已清仓则留现金；'
              '调仓份额为正表示买入，为负表示卖出。')

    ok = _push_table(f'本月调仓提醒 · {month_key}', intro, table_rows, footer)

    # 推送成功后更新份额台账（下次月度调仓以此为“当前份额”）
    if ok and new_shares:
        for c, sh in new_shares.items():
            funds[c]['shares'] = sh
            funds[c]['rebalanced'] = str(data_end)
        with open(sf, 'w', encoding='utf-8') as f:
            json.dump(mstate, f, ensure_ascii=False, indent=2)

    state.update({'last_month': month_key, 'weights': {c: float(w) for c, w in W.items()},
                  'data_end': str(data_end), 'pushed_at': datetime.datetime.now().strftime('%Y-%m-%d %H:%M')})
    with open(STATE_FILE, 'w', encoding='utf-8') as f:
        json.dump(state, f, ensure_ascii=False, indent=2)
    log('已推送' if ok else '推送未成功（webhook 未配置？），本月将不再重复提醒')


def _push_table(title, intro, rows, footer):
    try:
        import importlib
        return importlib.import_module('推送').push_table_card(
            title, intro=intro,
            columns=[('fund', '基金', 'auto'), ('wt', '目标权重', 'auto'),
                     ('nav', '净值', 'auto'), ('shares', '当前份额', 'auto'),
                     ('amount', '当前金额', 'auto'), ('dshares', '调仓份额', 'auto'),
                     ('damount', '调仓金额', 'auto')],
            rows=rows, footer=footer, template='blue')
    except Exception as e:
        log(f'推送模块异常：{e}')
        return False


def _push_card(title, blocks, template='blue'):
    try:
        import importlib
        return importlib.import_module('推送').push_card(title, blocks, template)
    except Exception as e:
        log(f'推送模块异常：{e}')
        return False


def _push(text):
    try:
        import importlib
        return importlib.import_module('推送').push(text)
    except Exception as e:
        log(f'推送模块异常：{e}')
        return False


def _configured():
    try:
        import importlib
        mod = importlib.import_module('推送')
        return bool((mod._cfg().get('webhook') or '').strip())
    except Exception:
        return False


def _push_fail(script):
    _push(f'【组合投资】月度调仓提醒未完成：{script} 执行失败，请手动检查。')


if __name__ == '__main__':
    main()
