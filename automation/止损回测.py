# -*- coding: utf-8 -*-
"""单基金止损回测（月度重估 + 固定份额 + 日度止损）：2026-01-01 至今。

口径：
  每月首个交易日用当时可得数据重估最大夏普权重并调回，两次调仓之间份额固定；
  日度检查止损：权益自高点回撤 ≥10% 减至 70%、≥18% 减至 40%、≥25% 清仓；债券 ≥3% 清仓；
  权益止损腾出资金优先买债券，债券自身已清仓时转现金（收益按 0 计）；
  冷静期 N 个交易日内禁止回补；回补需近 20 日累计收益 > 0 且回撤修复至 -5% 以内。
"""
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
import glob
import warnings

from 止损引擎 import simulate, BOND

warnings.filterwarnings('ignore')

plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

base_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(base_dir)
# 优先使用自动抓取的最新数据文件，没有则回退到手工维护的 .xls
auto_files = sorted(glob.glob(os.path.join(base_dir, "每日基金净值_自动_*.xlsx")))
file_path = auto_files[-1] if auto_files else os.path.join(base_dir, "每日基金净值.xls")
print(f"数据文件：{os.path.basename(file_path)}")
out = base_dir

df_raw = pd.read_excel(file_path, sheet_name="区间回报", index_col=0)
df_raw = df_raw[~df_raw.index.astype(str).str.contains('数据来源')]
df_raw.index = pd.to_datetime(df_raw.index, errors='coerce')
df_raw = df_raw[df_raw.index.notna()]
df_raw.sort_index(inplace=True)
df_raw = df_raw.dropna(axis=1, how='all').apply(pd.to_numeric, errors='coerce')
df_raw = df_raw.ffill()
df_raw = df_raw.loc[:, df_raw.notna().sum() >= 260]   # 面板对齐：按各基金成立日进入
df_cum = df_raw / 100.0
prices = 1 + df_cum

BT_START, RFR = '2026-01-01', 0.02
bt_dates = df_cum.index[df_cum.index >= BT_START]
print(f"\n回测区间 {bt_dates[0].date()} ~ {bt_dates[-1].date()}，共 {len(bt_dates)} 个交易日")


def metrics(nav):
    r = nav.pct_change().dropna()
    cum = nav.iloc[-1] - 1
    ann = (1 + cum) ** (252 / len(r)) - 1
    vol = r.std() * np.sqrt(252)
    dd = (nav / nav.cummax() - 1).min()
    return {'累计收益': cum, '年化收益': ann, '年化波动': vol,
            '夏普': (ann - RFR) / vol if vol > 0 else np.nan,
            '最大回撤': dd, 'Calmar': ann / abs(dd) if dd < 0 else np.nan}


def count_whipsaw(ev, days=14):
    """动作（清仓/减半）后 days 个自然日内又回补 = 打脸"""
    if ev.empty:
        return 0
    n = 0
    for c, g in ev.groupby('基金'):
        g = g.sort_values('日期').reset_index(drop=True)
        for i in range(len(g) - 1):
            if g.loc[i, '事件'] != '回补' and g.loc[i + 1, '事件'] == '回补':
                if (g.loc[i + 1, '日期'] - g.loc[i, '日期']).days <= days:
                    n += 1
    return n


CACHE = {}
LEVELS = [(0.10, 0.70), (0.18, 0.40), (0.25, 0.00)]   # 权益阶梯减仓
base = simulate(prices, bt_dates[0], stop=False, cache=CACHE)
COOLDOWNS = [0, 5, 10, 20, 30]
sims = {cd: simulate(prices, bt_dates[0], stop=True, cooldown=cd, levels=LEVELS, cache=CACHE)
        for cd in COOLDOWNS}


def rule_label(cd):
    return f'阶梯止损 10%/18%/25% + 冷静 {cd} 日'


rows = {'无止损': metrics(base['nav'])}
for cd in COOLDOWNS:
    sim = sims[cd]
    m = metrics(sim['nav'])
    ev = sim['events']
    m['止损事件数'] = int((ev['事件'] != '回补').sum()) if not ev.empty else 0
    m['打脸次数(≤10日回补)'] = count_whipsaw(ev)
    m['日均非满仓只数'] = sim['half'].mean()
    m['债券权重峰值'] = sim['bond_w'].max()
    m['现金权重峰值'] = sim['cash_w'].max()
    rows[rule_label(cd)] = m
cmp_df = pd.DataFrame(rows).T
print("\n===== 回测指标（月度重估 + 固定份额 + 日度止损，2026-01-01 至今） =====")
print(cmp_df.astype(float).map(lambda x: f'{x:.4f}' if abs(x) < 2 else f'{x:.2f}'))
print("\n===== 冷静期对比（核心指标） =====")
core = cmp_df[['累计收益', '年化波动', '夏普', '最大回撤', '止损事件数', '打脸次数(≤10日回补)']]
print(core.map(lambda x: f'{x:.4f}' if abs(x) < 2 else f'{x:.0f}'))

MAIN_CD = 20
main = sims[MAIN_CD]
nav_base = base['nav']
nav_main = main['nav']
half_main = main['half']
bond_main = main['bond_w']
cash_main = main['cash_w']
ev_main = main['events']

fig, ax = plt.subplots(figsize=(12, 6))
ax.plot(bt_dates, (nav_base - 1) * 100, label=f'无止损（{nav_base.iloc[-1] - 1:.2%}）',
        color='#4C72B0', lw=2)
for cd, c in [(0, '#DD8452'), (MAIN_CD, '#55A868')]:
    nv = sims[cd]['nav']
    ax.plot(bt_dates, (nv - 1) * 100, label=f'阶梯止损10%/18%/25%+冷静{cd}日（{nv.iloc[-1] - 1:.2%}）',
            color=c, lw=2)
ax.set_title('单基金止损回测：冷静期对比（累计收益率，月度重估+固定份额）')
ax.set_xlabel('日期'); ax.set_ylabel('累计收益率（%）')
ax.legend(); ax.grid(linestyle='--', alpha=0.5)
plt.tight_layout(); plt.savefig(os.path.join(out, '单基金止损_累计收益对比.png'), dpi=300); plt.close()

fig, ax = plt.subplots(figsize=(12, 6))
for nv, lb, c in [(nav_base, '无止损', '#4C72B0'), (sims[0]['nav'], '阶梯止损10%/18%/25%+冷静0日', '#DD8452'),
                  (nav_main, f'阶梯止损10%/18%/25%+冷静{MAIN_CD}日', '#55A868')]:
    dd = (nv / nv.cummax() - 1) * 100
    ax.plot(bt_dates, dd, label=f'{lb}（最大回撤 {dd.min():.2f}%）', color=c, lw=2)
ax.set_title('单基金止损回测：回撤对比')
ax.set_xlabel('日期'); ax.set_ylabel('回撤（%）')
ax.legend(); ax.grid(linestyle='--', alpha=0.5)
plt.tight_layout(); plt.savefig(os.path.join(out, '单基金止损_回撤对比.png'), dpi=300); plt.close()

fig, ax1 = plt.subplots(figsize=(12, 5))
ax1.step(bt_dates, half_main.values, where='post', color='#55A868', lw=1.8, label='减仓未回补基金只数')
ax1.set_ylabel('减仓未回补只数', color='#55A868')
ax2 = ax1.twinx()
ax2.plot(bt_dates, bond_main.values * 100, color='#DD8452', lw=1.5, label='债券基金权重（%）')
ax2.plot(bt_dates, cash_main.values * 100, color='#8172B3', lw=1.5, linestyle='--', label='现金权重（%）')
ax2.set_ylabel('权重（%）', color='#DD8452')
lines = ax1.get_lines() + ax2.get_lines()
ax1.legend(lines, [l.get_label() for l in lines], loc='upper left')
ax1.set_title(f'单基金止损（冷静{MAIN_CD}日）：减仓未回补只数与资金腾挪')
ax1.grid(linestyle='--', alpha=0.5)
plt.tight_layout(); plt.savefig(os.path.join(out, '单基金止损_状态.png'), dpi=300); plt.close()

cd_df = cmp_df.loc[[rule_label(cd) for cd in COOLDOWNS]]
fig, axes = plt.subplots(1, 3, figsize=(16, 4.8))
axes[0].bar([f'{cd}日' for cd in COOLDOWNS], cd_df['累计收益'] * 100, color='#55A868')
axes[0].set_title('累计收益率（%）')
axes[1].bar([f'{cd}日' for cd in COOLDOWNS], -cd_df['最大回撤'] * 100, color='#c44e52')
axes[1].set_title('最大回撤（%）')
axes[2].bar([f'{cd}日' for cd in COOLDOWNS], cd_df['打脸次数(≤10日回补)'], color='#4C72B0')
axes[2].set_title('打脸次数（≤10日回补）')
for a in axes:
    a.set_xlabel('冷静期长度'); a.grid(axis='y', linestyle='--', alpha=0.5)
plt.tight_layout(); plt.savefig(os.path.join(out, '单基金止损_冷静期对比.png'), dpi=300); plt.close()

xlsx = os.path.join(out, '单基金止损回测结果.xlsx')
with pd.ExcelWriter(xlsx, engine='openpyxl') as w:
    cmp_df.to_excel(w, sheet_name='指标对比_冷静期')
    base['weights'].iloc[0].to_frame('起始权重').to_excel(w, sheet_name='配置权重')
    daily = pd.DataFrame({
        '无止损_净值': nav_base, '无止损_累计收益': nav_base - 1,
        '无止损_回撤': nav_base / nav_base.cummax() - 1,
        f'冷静{MAIN_CD}日_净值': nav_main, f'冷静{MAIN_CD}日_累计收益': nav_main - 1,
        f'冷静{MAIN_CD}日_回撤': nav_main / nav_main.cummax() - 1,
        '半仓基金只数': half_main, '债券基金权重': bond_main, '现金权重': cash_main,
    })
    daily.to_excel(w, sheet_name=f'每日明细_冷静{MAIN_CD}日')
    if not ev_main.empty:
        ev_main.to_excel(w, sheet_name=f'止损事件_冷静{MAIN_CD}日', index=False)
    for cd in COOLDOWNS:
        ev = sims[cd]['events']
        if not ev.empty:
            ev.to_excel(w, sheet_name=f'止损事件_冷静{cd}', index=False)
    mo = pd.DataFrame({'无止损': nav_base.resample('ME').last().pct_change(),
                       '冷静0日': sims[0]['nav'].resample('ME').last().pct_change(),
                       f'冷静{MAIN_CD}日': nav_main.resample('ME').last().pct_change()})
    mo.to_excel(w, sheet_name='月度收益')

print(f"\n结果已保存至 {out}")
