import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from scipy.optimize import minimize
import datetime
import os
import glob
import warnings
warnings.filterwarnings('ignore')

plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

base_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(base_dir)
# 优先使用自动抓取的数据文件（取日期最新的一个），没有则回退到手工维护的 .xls
auto_files = sorted(glob.glob(os.path.join(base_dir, "每日基金净值_自动_*.xlsx")))
file_path = auto_files[-1] if auto_files else os.path.join(base_dir, "每日基金净值.xls")
print(f"数据文件：{os.path.basename(file_path)}")
output_dir = base_dir
os.makedirs(output_dir, exist_ok=True)

df_raw = pd.read_excel(file_path, sheet_name="区间回报", index_col=0)
df_raw = df_raw[~df_raw.index.astype(str).str.contains('数据来源')]
df_raw.index = pd.to_datetime(df_raw.index, errors='coerce')
df_raw = df_raw[df_raw.index.notna()]
df_raw.sort_index(inplace=True)
df_raw = df_raw.dropna(axis=1, how='all')
df_raw = df_raw.apply(pd.to_numeric, errors='coerce')
# 004253 仅用于止损监控，不进 MV/RP 优化池（12 月起由 008143 替换）
df_raw = df_raw.drop(columns=['国泰黄金ETF联接C'], errors='ignore')
# 逐基金保留各自起点，不再整表 dropna（否则会被最晚成立的基金截断）
df_raw = df_raw.ffill()
df_raw = df_raw.loc[df_raw.notna().any(axis=1)]
df_cum = df_raw / 100.0

print(f"数据日期范围：{df_cum.index.min()} 至 {df_cum.index.max()}")
print(f"基金数量：{df_cum.shape[1]}")

daily_ret = df_cum.diff().iloc[1:]

def negative_sharpe(weights, mean_ret, cov_matrix, risk_free=0.02):
    port_ret = np.sum(weights * mean_ret)
    port_vol = np.sqrt(weights.T @ cov_matrix @ weights)
    if port_vol == 0:
        return 1e10
    return -(port_ret - risk_free) / port_vol

def optimize_sharpe_multistart(mean_ret, cov_matrix, n_starts=100, max_weight=0.3):
    """最大夏普：凸化精确求解（n_starts 仅为兼容旧调用保留，内部不再用随机起点）。

    令 y = w / (mu·w - rf)，则 max (mu·w-rf)/sqrt(w'Σw) 等价于凸二次规划：
        min y'Σy   s.t. (mu-rf)'y = 1, y >= 0, y_i <= max_weight * sum(y)
    凸问题的局部最优即全局最优，因此 w = y / sum(y) 为精确解。
    """
    mu = np.asarray(mean_ret, dtype=float)
    cov = np.asarray(cov_matrix, dtype=float)
    n = len(mu)
    ex = mu - 0.02

    def sharpe(w):
        return (w @ mu - 0.02) / np.sqrt(w @ cov @ w)

    # 无正超额收益时最大夏普无解，退化为最小方差
    if np.max(ex) <= 0:
        res = minimize(lambda w: w @ cov @ w, np.full(n, 1 / n), method='SLSQP',
                       bounds=[(0, max_weight)] * n,
                       constraints=({'type': 'eq', 'fun': lambda w: np.sum(w) - 1},),
                       options={'ftol': 1e-14, 'maxiter': 2000, 'disp': False})
        w = res.x if res.success else np.full(n, 1 / n)
        return w, sharpe(w)

    y0 = np.clip(ex, 0, None)
    y0 = y0 / (ex @ y0)
    constraints = ({'type': 'eq', 'fun': lambda y: ex @ y - 1},
                   {'type': 'ineq', 'fun': lambda y: max_weight * np.sum(y) - y})
    res = minimize(lambda y: y @ cov @ y, y0, method='SLSQP',
                   bounds=[(0, None)] * n, constraints=constraints,
                   options={'ftol': 1e-14, 'maxiter': 2000, 'disp': False})
    if not res.success or np.sum(res.x) <= 0:
        return None, -np.inf
    w = res.x / np.sum(res.x)
    return w, sharpe(w)

def optimize_risk_parity(cov_matrix):
    """风险平价（等风险贡献 ERC）：坐标下降精确求解，无随机起点，结果稳定。"""
    cov = np.asarray(cov_matrix, dtype=float)
    n = cov.shape[0]
    diag = np.diag(cov).copy()
    w = 1.0 / np.sqrt(diag)
    for _ in range(5000):
        w_new = w.copy()
        for i in range(n):
            b = float(cov[i] @ w - cov[i, i] * w[i])
            disc = b * b + 4 * diag[i] / n
            w_new[i] = (-b + np.sqrt(disc)) / (2 * diag[i]) if disc > 0 else w[i]
        if np.max(np.abs(w_new - w)) < 1e-15:
            w = w_new
            break
        w = w_new
    return w / w.sum()

one_year_ago = df_cum.index.max() - pd.DateOffset(years=1)
daily_ret_1y = daily_ret[daily_ret.index >= one_year_ago]
corr_1y = daily_ret_1y.corr()

three_years_ago = df_cum.index.max() - pd.DateOffset(years=3)
df_cum_3y = df_cum[df_cum.index >= three_years_ago]
try:
    monthly_last = df_cum_3y.resample('ME').last()
except ValueError:
    monthly_last = df_cum_3y.resample('M').last()
monthly_ret = monthly_last.diff().dropna(how='all')
# 期望收益按各基金实际可得月份计算，要求至少 24 个月，避免短样本噪声
short = [c for c in monthly_ret.columns if monthly_ret[c].count() < 24]
if short:
    print(f"样本不足 24 个月，不参与优化：{'、'.join(short)}")
monthly_ret = monthly_ret.drop(columns=short, errors='ignore')
monthly_ret = monthly_ret.dropna(how='all', axis=0).dropna(how='all', axis=1)

avg_monthly_ret = monthly_ret.mean(axis=0).sort_values(ascending=False)
plt.figure(figsize=(12, 6))
ax = avg_monthly_ret.plot(kind='bar', color='skyblue')
plt.title('近三年各基金平均月收益率（含填充）')
plt.xlabel('基金')
plt.ylabel('平均月收益率')
plt.xticks(rotation=45, ha='right')
for i, v in enumerate(avg_monthly_ret):
    ax.text(i, v + 0.001, f'{v:.2%}', ha='center', va='bottom', fontsize=8)
plt.tight_layout()
plt.savefig(os.path.join(output_dir, '各基金平均月收益率.png'), dpi=300)
plt.close()

monthly_avg_overall = monthly_ret.mean(axis=1)
plt.figure(figsize=(14, 6))
plt.plot(monthly_avg_overall.index, monthly_avg_overall.values, marker='o', linestyle='-')
for x, y in zip(monthly_avg_overall.index, monthly_avg_overall.values):
    plt.text(x, y + 0.002, f'{y:.2%}', ha='center', va='bottom', fontsize=7, rotation=45)
plt.title('近三年月度整体平均收益率')
plt.xlabel('月份')
plt.ylabel('月度收益率')
plt.grid(True, linestyle='--', alpha=0.7)
plt.tight_layout()
plt.savefig(os.path.join(output_dir, '月度整体平均收益率走势.png'), dpi=300)
plt.close()

monthly_ret_display = monthly_ret.copy()
monthly_ret_display.index = monthly_ret_display.index.strftime('%Y-%m')
fig, ax = plt.subplots(figsize=(18, 8))
sns.heatmap(monthly_ret_display.T, annot=True, fmt='.2%', cmap='RdBu_r', center=0, ax=ax,
            cbar_kws={'label': '月度收益率'}, annot_kws={'size': 7})
ax.set_title('近三年各基金月度收益率热力图')
ax.set_xlabel('月份')
ax.set_ylabel('基金名称')
plt.xticks(rotation=45, ha='right')
plt.yticks(rotation=0)
plt.tight_layout()
plt.savefig(os.path.join(output_dir, '月度收益率热力图.png'), dpi=300)
plt.close()

monthly_ret_3y = monthly_ret.copy()
monthly_ret_3y = monthly_ret_3y.loc[:, monthly_ret_3y.std() > 1e-8]
mean_ret_annual = monthly_ret_3y.mean() * 12

daily_ret_1y = daily_ret[daily_ret.index >= one_year_ago]
daily_ret_1y = daily_ret_1y.loc[:, daily_ret_1y.std() > 1e-8]
cov_1y = daily_ret_1y.cov() * 252

common_cols = [col for col in monthly_ret_3y.columns if col in daily_ret_1y.columns]
mean_ret_annual = mean_ret_annual[common_cols]
cov_matrix = cov_1y.loc[common_cols, common_cols]

epsilon = 1e-6
cov_matrix_reg = cov_matrix + np.eye(cov_matrix.shape[0]) * epsilon

opt_weights, max_sharpe = optimize_sharpe_multistart(mean_ret_annual.values, cov_matrix_reg.values, n_starts=100, max_weight=0.3)
if opt_weights is None:
    print("30%上限无解，放宽至100%")
    opt_weights, max_sharpe = optimize_sharpe_multistart(mean_ret_annual.values, cov_matrix_reg.values, n_starts=100, max_weight=1.0)
    if opt_weights is None:
        raise RuntimeError("优化失败")
weights_series = pd.Series(opt_weights, index=common_cols, dtype=float)
weights_series_nonzero = weights_series[weights_series > 1e-6]
print("\n最优配置权重：")
print(weights_series_nonzero.sort_values(ascending=False))

# === 相关系数矩阵（高亮最优配置基金） ===
selected_funds = weights_series_nonzero.index.tolist()
all_funds = list(corr_1y.columns)
other_funds = [f for f in all_funds if f not in selected_funds]
ordered_funds = selected_funds + other_funds
corr_ordered = corr_1y.loc[ordered_funds, ordered_funds]

fig, ax = plt.subplots(figsize=(14, 10))
sns.heatmap(corr_ordered, annot=True, fmt='.2f', cmap='coolwarm',
            cbar_kws={'label': '相关系数'}, ax=ax,
            linewidths=0.5, linecolor='gray')
n_selected = len(selected_funds)
rect = plt.Rectangle((0, 0), n_selected, n_selected,
                     linewidth=3, edgecolor='lime', facecolor='none', zorder=10)
ax.add_patch(rect)
for i, label in enumerate(ax.get_xticklabels()):
    if i < n_selected:
        label.set_color('red')
        label.set_fontweight('bold')
for i, label in enumerate(ax.get_yticklabels()):
    if i < n_selected:
        label.set_color('red')
        label.set_fontweight('bold')
plt.title('近一年基金日收益率相关系数矩阵（绿框内为最优配置基金）')
plt.tight_layout()
plt.savefig(os.path.join(output_dir, '相关系数矩阵.png'), dpi=300)
plt.close()
print("已保存相关系数矩阵图（高亮最优配置基金）")

port_ret = np.sum(weights_series * mean_ret_annual)
port_vol = np.sqrt(weights_series @ cov_matrix_reg @ weights_series)
print(f"组合年化收益率：{port_ret:.4%}，年化波动率：{port_vol:.4%}，夏普比率：{max_sharpe:.4f}")

plt.figure(figsize=(10, 6))
ax = weights_series_nonzero.sort_values().plot(kind='barh', color='lightgreen')
plt.title('最优投资组合权重')
plt.xlabel('权重')
for i, v in enumerate(weights_series_nonzero.sort_values()):
    ax.text(v + 0.002, i, f'{v:.2%}', va='center')
plt.tight_layout()
plt.savefig(os.path.join(output_dir, '最优权重.png'), dpi=300)
plt.close()

full_weights = pd.Series(0.0, index=df_cum.columns)
full_weights[weights_series.index] = weights_series.values
port_cum_abs = (df_cum * full_weights).sum(axis=1, min_count=1)

today = port_cum_abs.index.max()
data_start = port_cum_abs.index.min()
start_ytd = max(datetime.datetime(2026, 1, 1), data_start)
start_3m = max(today - pd.DateOffset(months=3), data_start)
start_1m = max(today - pd.DateOffset(months=1), data_start)

def plot_interval_cumulative(port_series, start_date, title, filename):
    sub_abs = port_series[port_series.index >= start_date]
    if sub_abs.empty:
        return None
    sub_rel = sub_abs - sub_abs.iloc[0]
    plt.figure(figsize=(12, 6))
    plt.plot(sub_rel.index, sub_rel.values, linewidth=2, color='green')
    plt.text(sub_rel.index[0], sub_rel.iloc[0], f'起始:{sub_rel.iloc[0]:.2%}', ha='center', va='bottom', fontsize=9)
    plt.text(sub_rel.index[-1], sub_rel.iloc[-1], f'终值:{sub_rel.iloc[-1]:.2%}', ha='center', va='bottom', fontsize=9)
    plt.title(title)
    plt.xlabel('日期')
    plt.ylabel('区间累计收益率（起点0%）')
    plt.grid(True, linestyle='--', alpha=0.6)
    y_min = min(0, sub_rel.min())
    y_max = max(sub_rel.max(), 0.01)
    plt.ylim(bottom=y_min, top=y_max*1.1)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, filename), dpi=300)
    plt.close()
    return sub_rel

sub_ytd = plot_interval_cumulative(port_cum_abs, start_ytd, '今年以来累计收益率（起点0%）', '回测_今年以来_累计收益率.png')
sub_3m = plot_interval_cumulative(port_cum_abs, start_3m, '最近三个月累计收益率（起点0%）', '回测_最近三个月_累计收益率.png')
sub_1m = plot_interval_cumulative(port_cum_abs, start_1m, '最近一个月累计收益率（起点0%）', '回测_最近一个月_累计收益率.png')

# ==================== 风险平价模型（波动率取最近一年） ====================
# 精确 ERC 求解，不设单只权重上限
rp_opt_weights = optimize_risk_parity(cov_matrix_reg.values)
if rp_opt_weights is None:
    raise RuntimeError("风险平价优化失败")
rp_weights_series = pd.Series(rp_opt_weights, index=common_cols, dtype=float)
rp_weights_nonzero = rp_weights_series[rp_weights_series > 1e-6]
print("\n风险平价配置权重：")
print(rp_weights_nonzero.sort_values(ascending=False))

rp_port_ret = np.sum(rp_weights_series * mean_ret_annual)
rp_port_vol = np.sqrt(rp_weights_series @ cov_matrix_reg @ rp_weights_series)
rp_sharpe = (rp_port_ret - 0.02) / rp_port_vol if rp_port_vol > 0 else np.nan
print(f"风险平价组合：年化收益率 {rp_port_ret:.4%}，年化波动率 {rp_port_vol:.4%}，夏普比率 {rp_sharpe:.4f}")

# 各资产风险贡献占比（风险平价核心指标）
rp_sigma = np.sqrt(rp_weights_series @ cov_matrix_reg @ rp_weights_series)
rp_mrc = cov_matrix_reg @ rp_weights_series
rp_rc = rp_weights_series * rp_mrc / rp_sigma
rp_rc_pct = (rp_rc / rp_sigma).sort_values(ascending=False)
print("\n风险平价风险贡献占比：")
print(rp_rc_pct)

# 风险平价权重图
plt.figure(figsize=(10, 6))
ax = rp_weights_nonzero.sort_values().plot(kind='barh', color='cornflowerblue')
plt.title('风险平价投资组合权重（波动率取近一年）')
plt.xlabel('权重')
for i, v in enumerate(rp_weights_nonzero.sort_values()):
    ax.text(v + 0.002, i, f'{v:.2%}', va='center')
plt.tight_layout()
plt.savefig(os.path.join(output_dir, '风险平价_最优权重.png'), dpi=300)
plt.close()

# 风险贡献占比图
plt.figure(figsize=(10, 6))
ax = rp_rc_pct.sort_values().plot(kind='barh', color='salmon')
plt.title('风险平价各资产风险贡献占比')
plt.xlabel('风险贡献占比')
for i, v in enumerate(rp_rc_pct.sort_values()):
    ax.text(v + 0.002, i, f'{v:.2%}', va='center')
plt.tight_layout()
plt.savefig(os.path.join(output_dir, '风险平价_风险贡献.png'), dpi=300)
plt.close()

# 两模型权重对比图
compare_df = pd.DataFrame({'均值方差': weights_series, '风险平价': rp_weights_series})
compare_df = compare_df[compare_df.sum(axis=1) > 1e-6]
compare_df = compare_df.sort_values('风险平价')
ax = compare_df.plot(kind='barh', figsize=(12, 8), color=['#4C72B0', '#DD8452'])
plt.title('均值方差 vs 风险平价 权重对比')
plt.xlabel('权重')
plt.ylabel('基金名称')
plt.tight_layout()
plt.savefig(os.path.join(output_dir, '两模型权重对比.png'), dpi=300)
plt.close()

# 风险平价组合净值与回测
rp_full_weights = pd.Series(0.0, index=df_cum.columns)
rp_full_weights[rp_weights_series.index] = rp_weights_series.values
rp_port_cum_abs = (df_cum * rp_full_weights).sum(axis=1, min_count=1)

rp_sub_ytd = plot_interval_cumulative(rp_port_cum_abs, start_ytd, '风险平价 今年以来累计收益率（起点0%）', '风险平价_回测_今年以来_累计收益率.png')
rp_sub_3m = plot_interval_cumulative(rp_port_cum_abs, start_3m, '风险平价 最近三个月累计收益率（起点0%）', '风险平价_回测_最近三个月_累计收益率.png')
rp_sub_1m = plot_interval_cumulative(rp_port_cum_abs, start_1m, '风险平价 最近一个月累计收益率（起点0%）', '风险平价_回测_最近一个月_累计收益率.png')

# 两模型累计收益对比（近一年）
comp_start = max(one_year_ago, data_start)
sub_mv_cmp = port_cum_abs[port_cum_abs.index >= comp_start]
sub_rp_cmp = rp_port_cum_abs[rp_port_cum_abs.index >= comp_start]
rel_mv_cmp = sub_mv_cmp - sub_mv_cmp.iloc[0]
rel_rp_cmp = sub_rp_cmp - sub_rp_cmp.iloc[0]
plt.figure(figsize=(13, 6))
plt.plot(rel_mv_cmp.index, rel_mv_cmp.values, label='均值方差', color='#4C72B0', linewidth=2)
plt.plot(rel_rp_cmp.index, rel_rp_cmp.values, label='风险平价', color='#DD8452', linewidth=2)
plt.title('两模型累计收益率对比（近一年，起点0%）')
plt.xlabel('日期')
plt.ylabel('区间累计收益率')
plt.legend()
plt.grid(True, linestyle='--', alpha=0.6)
plt.tight_layout()
plt.savefig(os.path.join(output_dir, '两模型累计收益对比.png'), dpi=300)
plt.close()

# 两模型回测合并对比（今年以来 / 最近三个月 / 最近一个月）
def _rel(series, start):
    sub = series[series.index >= start]
    return sub - sub.iloc[0] if not sub.empty else sub

_periods = [('今年以来', start_ytd), ('最近三个月', start_3m), ('最近一个月', start_1m)]
fig, axes = plt.subplots(1, 3, figsize=(20, 6))
for ax, (name, start) in zip(axes, _periods):
    mv = _rel(port_cum_abs, start)
    rp = _rel(rp_port_cum_abs, start)
    ax.plot(mv.index, mv.values, label='均值方差', color='#4C72B0', linewidth=2)
    ax.plot(rp.index, rp.values, label='风险平价', color='#DD8452', linewidth=2)
    ax.set_title(f'{name}累计收益率（起点0%）')
    ax.set_xlabel('日期')
    ax.set_ylabel('区间累计收益率')
    ax.grid(True, linestyle='--', alpha=0.6)
    ax.legend()
    for s, c in [(mv, '#4C72B0'), (rp, '#DD8452')]:
        if not s.empty:
            ax.text(s.index[-1], s.iloc[-1], f'{s.iloc[-1]:.2%}', ha='right', va='bottom', fontsize=9, color=c)
plt.tight_layout()
plt.savefig(os.path.join(output_dir, '两模型回测对比_三区间.png'), dpi=300)
plt.close()

# 两模型回撤对比（近一年）
def _drawdown(series):
    nav = 1 + series                      # 累计收益率转净值
    return nav / nav.cummax() - 1          # 相对历史最高净值的回撤

print(f"诊断：近一年累计收益率区间 均值方差 [{port_cum_abs[port_cum_abs.index >= comp_start].min():.2%}, {port_cum_abs[port_cum_abs.index >= comp_start].max():.2%}]，"
      f"风险平价 [{rp_port_cum_abs[rp_port_cum_abs.index >= comp_start].min():.2%}, {rp_port_cum_abs[rp_port_cum_abs.index >= comp_start].max():.2%}]")
mv_dd = _drawdown(port_cum_abs)
rp_dd = _drawdown(rp_port_cum_abs)
mv_dd = mv_dd[mv_dd.index >= comp_start]
rp_dd = rp_dd[rp_dd.index >= comp_start]
print(f"\n近一年最大回撤：均值方差 {mv_dd.min():.2%}，风险平价 {rp_dd.min():.2%}")

plt.figure(figsize=(13, 6))
plt.plot(mv_dd.index, mv_dd.values, label=f"均值方差（最大回撤 {mv_dd.min():.2%}）", color='#4C72B0', linewidth=2)
plt.plot(rp_dd.index, rp_dd.values, label=f"风险平价（最大回撤 {rp_dd.min():.2%}）", color='#DD8452', linewidth=2)
plt.title('两模型回撤对比（近一年）')
plt.xlabel('日期')
plt.ylabel('回撤')
plt.legend()
plt.grid(True, linestyle='--', alpha=0.6)
plt.tight_layout()
plt.savefig(os.path.join(output_dir, '两模型回撤对比.png'), dpi=300)
plt.close()

dd_df = pd.DataFrame({'均值方差_回撤': mv_dd, '风险平价_回撤': rp_dd})

excel_path = os.path.join(output_dir, '投资组合分析明细.xlsx')
with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
    df_raw.to_excel(writer, sheet_name='原始累计收益率_百分比')
    df_cum.to_excel(writer, sheet_name='累计收益率_小数')
    daily_ret.to_excel(writer, sheet_name='日收益率')
    cov_matrix.to_excel(writer, sheet_name='年化协方差矩阵_日度近一年')
    corr_1y.to_excel(writer, sheet_name='相关系数矩阵_日度近一年')
    monthly_ret.to_excel(writer, sheet_name='近三年月收益率')
    avg_monthly_ret.to_frame(name='平均月收益率').to_excel(writer, sheet_name='基金月均收益率')
    monthly_avg_overall.to_frame(name='月度整体平均收益率').to_excel(writer, sheet_name='月度整体平均收益率')
    weights_series_nonzero.to_frame(name='权重').to_excel(writer, sheet_name='最优配置权重')
    port_cum_abs.to_frame(name='绝对累计收益率').to_excel(writer, sheet_name='组合绝对累计收益率')
    if sub_ytd is not None:
        sub_ytd.to_frame(name='相对累计收益率').to_excel(writer, sheet_name='回测_今年以来_起点0')
    if sub_3m is not None:
        sub_3m.to_frame(name='相对累计收益率').to_excel(writer, sheet_name='回测_最近三个月_起点0')
    if sub_1m is not None:
        sub_1m.to_frame(name='相对累计收益率').to_excel(writer, sheet_name='回测_最近一个月_起点0')
    mean_ret_annual.to_frame(name='年化收益率_月度近三年').to_excel(writer, sheet_name='年化收益率_月度近三年')
    pd.DataFrame({
        '指标': ['组合年化收益率', '组合年化波动率', '夏普比率', '基金数量', '数据起始日期', '数据结束日期'],
        '数值': [f'{port_ret:.4%}', f'{port_vol:.4%}', f'{max_sharpe:.4f}',
                 len(common_cols), df_cum.index.min().strftime('%Y-%m-%d'), df_cum.index.max().strftime('%Y-%m-%d')]
    }).to_excel(writer, sheet_name='组合统计概览', index=False)

    # ===== 风险平价模型输出 =====
    rp_weights_nonzero.to_frame(name='权重').to_excel(writer, sheet_name='风险平价配置权重')
    rp_rc_pct.to_frame(name='风险贡献占比').to_excel(writer, sheet_name='风险平价风险贡献')
    rp_port_cum_abs.to_frame(name='绝对累计收益率').to_excel(writer, sheet_name='风险平价组合累计收益率')
    if rp_sub_ytd is not None:
        rp_sub_ytd.to_frame(name='相对累计收益率').to_excel(writer, sheet_name='风险平价_回测_今年以来_起点0')
    if rp_sub_3m is not None:
        rp_sub_3m.to_frame(name='相对累计收益率').to_excel(writer, sheet_name='风险平价_回测_最近三个月_起点0')
    if rp_sub_1m is not None:
        rp_sub_1m.to_frame(name='相对累计收益率').to_excel(writer, sheet_name='风险平价_回测_最近一个月_起点0')
    pd.DataFrame({
        '指标': ['组合年化收益率', '组合年化波动率', '夏普比率', '基金数量', '波动率区间'],
        '风险平价': [f'{rp_port_ret:.4%}', f'{rp_port_vol:.4%}', f'{rp_sharpe:.4f}',
                     len(common_cols), '近一年日度']
    }).to_excel(writer, sheet_name='风险平价统计概览', index=False)
    pd.DataFrame({
        '指标': ['年化收益率', '年化波动率', '夏普比率'],
        '均值方差': [f'{port_ret:.4%}', f'{port_vol:.4%}', f'{max_sharpe:.4f}'],
        '风险平价': [f'{rp_port_ret:.4%}', f'{rp_port_vol:.4%}', f'{rp_sharpe:.4f}']
    }).to_excel(writer, sheet_name='两模型对比', index=False)
    dd_df.to_excel(writer, sheet_name='两模型回撤对比_近一年')

print(f"\n所有图表和Excel文件已保存至：{output_dir}")
