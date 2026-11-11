# -*- coding: utf-8 -*-
"""自动获取 18 只基金每日净值（东方财富公开接口），生成与“每日基金净值.xls”同口径的表格。

数据源：http://fund.eastmoney.com/pingzhongdata/{code}.js
  - Data_ACWorthTrend（累计净值，含分红再投）→ 计算区间累计收益率，与原有 Excel 口径一致
  - Data_netWorthTrend（单位净值）→ 供核对
  注：接口时间戳为 UTC（16:00 UTC = 北京时间次日 00:00），需按 Asia/Shanghai 归一化，
      否则日期会整体早一天。
输出：每日基金净值_自动_YYYYMMDD.xlsx（sheet：区间回报 / 单位净值 / 说明）
可选：与同目录下 每日基金净值.xls 做一致性校验（中位偏差）。
用法：python 基金数据获取.py [起始日期，默认取现有 Excel 的首个日期]
"""
import requests, re, json, sys, time, os
import pandas as pd
import numpy as np

h = {'User-Agent': 'Mozilla/5.0', 'Referer': 'http://fund.eastmoney.com/'}
base_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(base_dir)

# 先读现有 Excel 的首个日期作为起算日（保证与新数据同口径）
fx = os.path.join(base_dir, '每日基金净值.xls')
chk = None
if os.path.exists(fx):
    chk = pd.read_excel(fx, sheet_name='区间回报', index_col=0)
    chk = chk[~chk.index.astype(str).str.contains('数据来源')]
    chk.index = pd.to_datetime(chk.index, errors='coerce')
    chk = chk[chk.index.notna()].sort_index()

START = sys.argv[1] if len(sys.argv) > 1 else (
    str(chk.index.min().date()) if chk is not None and len(chk) else '2023-11-01')

# 基金名称（与 Excel 列名一致）→ 基金代码（已用净值曲线逐一核对，含 A/C 份额）
FUNDS = {
    '国泰黄金ETF联接C': '004253',
    '嘉实中证稀有金属主题ETF发起联接C': '014111',
    '南方中证全指证券公司ETF联接C': '004070',
    '华泰保兴安悦债券C': '020741',
    '易方达恒生港股通红利低波动ETF联接发起式C': '021458',
    '富国中证煤炭指数C': '013275',
    '招商中证白酒指数C': '012414',
    '华泰柏瑞中证稀土产业ETF发起式联接C': '014332',
    '嘉实上证科创板芯片ETF发起联接C': '017470',
    '华宝中证沪港深新消费指数C': '017435',
    '易方达中证1000量化增强C': '017095',
    '永赢医药创新智选混合发起C': '015916',
    '东方新能源汽车主题混合': '400015',
    '华夏中证动漫游戏ETF发起式联接C': '012769',
    '招商中证机器人ETF发起联接C': '020482',
    '农银现代农业加混合': '001940',
    '富国中证军工指数(LOF)C': '013035',
    '天弘中证细分化工产业主题ETF联接C': '015897',
    '广发中证全指家用电器ETF联接C': '005064',
}

def _ts(ms):
    """东财时间戳（UTC）→ 北京时间交易日（自然日）"""
    return (pd.Timestamp(int(ms), unit='ms', tz='UTC')
            .tz_convert('Asia/Shanghai').normalize().tz_localize(None))

def fetch_js(code, field):
    r = requests.get(f'http://fund.eastmoney.com/pingzhongdata/{code}.js', headers=h, timeout=20)
    r.raise_for_status()
    m = re.search(field + r'\s*=\s*(\[.*?\]);', r.text)
    if not m:
        return None
    arr = json.loads(m.group(1))
    if field == 'Data_ACWorthTrend':
        s = pd.Series({_ts(p[0]): float(p[1]) for p in arr if p[1] is not None})
    else:
        s = pd.Series({_ts(p['x']): float(p['y']) for p in arr})
    s = s.sort_index()
    return s[~s.index.duplicated(keep='last')]

ac, nv = {}, {}
for name, code in FUNDS.items():
    for attempt in range(3):
        try:
            ac[name] = fetch_js(code, 'Data_ACWorthTrend')
            nv[name] = fetch_js(code, 'Data_netWorthTrend')
            break
        except Exception as e:
            print(f'  {name}({code}) 第{attempt+1}次失败: {e}')
            time.sleep(1 + attempt)
    if not ac.get(name) is not None:
        print(f'!! {name}({code}) 获取失败，请检查')
    else:
        print(f'{name}: {len(ac[name])} 条，{ac[name].index.min().date()} ~ {ac[name].index.max().date()}')
    time.sleep(0.15)

nav_df = pd.DataFrame(nv).sort_index()
ac_df = pd.DataFrame(ac).sort_index()
start_ts = pd.Timestamp(START)

# 复权净值：累计净值 - 单位净值 = 累计分红，其增量即当日分红额；
# 注意按“单只基金自己的日期序列”计算，避免因该基金当日无净值而丢掉收益
def _adjusted(nav_s, ac_s):
    if ac_s is None:
        return nav_s / nav_s.iloc[0]
    d = (ac_s.reindex(nav_s.index) - nav_s).ffill()
    div = d.diff().clip(lower=0).fillna(0)
    r = (nav_s + div) / nav_s.shift(1) - 1
    return (1 + r.fillna(0)).cumprod()

adj_df = pd.DataFrame({
    c: _adjusted(nav_df[c].dropna(), ac_df[c].dropna() if c in ac_df.columns else None)
    for c in nav_df.columns}).sort_index()

# 基准 = 起算日之前的最后一个复权净值（与现有 Excel 的区间回报同起点）
prev = adj_df[adj_df.index < start_ts]
base = prev.iloc[-1] if len(prev) else pd.Series(dtype=float, index=adj_df.columns)
adj_df = adj_df[adj_df.index >= start_ts]
nav_out = nav_df[nav_df.index >= start_ts]
ac_out = ac_df[ac_df.index >= start_ts]
base = base.reindex(adj_df.columns).fillna(adj_df.bfill().iloc[0])

# 区间累计收益率（%）= 复权净值 / 基准 - 1
cum_pct = (adj_df / base - 1) * 100
cum_pct = cum_pct.round(2)
cum_pct = cum_pct.dropna(how='all')

# 一致性校验：与同目录现有 Excel 比对（如存在）
if chk is not None:
    diffs = {}
    for c in cum_pct.columns:
        if c in chk.columns:
            a = cum_pct[c].dropna(); b = chk[c].dropna()
            both = pd.concat([a, b], axis=1, join='inner').dropna()
            diffs[c] = float(np.nanmedian(np.abs(both.iloc[:, 0] - both.iloc[:, 1])))
    print('\n=== 与现有 Excel 的一致性校验（中位绝对偏差，pct）===')
    for k, v in sorted(diffs.items(), key=lambda x: -x[1]):
        flag = 'OK' if v < 1.0 else '偏大，请复核'
        print(f'  {k}: {v:.3f}  {flag}')

stamp = pd.Timestamp.now().strftime('%Y%m%d')
out_path = os.path.join(base_dir, f'每日基金净值_自动_{stamp}.xlsx')
note = pd.DataFrame({'说明': [
    '数据来源：东方财富基金公开接口（pingzhongdata.js）',
    '区间回报 = 单位净值 / 起算日前一净值点 - 1，单位 %（未复权，与原有 Excel 口径一致）',
    '单位净值为原始单位净值；累计净值为 单位净值+累计分红，仅供核对',
    f'生成时间：{pd.Timestamp.now():%Y-%m-%d %H:%M:%S}；数据截止：{cum_pct.index.max().date()}',
    '基金代码映射见 基金数据获取.py 中 FUNDS 字典',
]})
with pd.ExcelWriter(out_path, engine='openpyxl') as w:
    cum_pct.to_excel(w, sheet_name='区间回报', index_label='日期')
    nav_out.to_excel(w, sheet_name='单位净值', index_label='日期')
    ac_out.to_excel(w, sheet_name='累计净值', index_label='日期')
    note.to_excel(w, sheet_name='说明', index=False)

print(f'\n已生成：{out_path}')
print(f'数据截止日期：{cum_pct.index.max().date()}，共 {len(cum_pct)} 行，{cum_pct.shape[1]} 列')
