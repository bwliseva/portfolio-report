// assets/charts.js — 组合投资报告图表
(function () {
  var D = window.RPT;
  var style = getComputedStyle(document.documentElement);
  var accent = style.getPropertyValue('--accent').trim();
  var accent2 = style.getPropertyValue('--accent2').trim();
  var ink = style.getPropertyValue('--ink').trim();
  var muted = style.getPropertyValue('--muted').trim();
  var rule = style.getPropertyValue('--rule').trim();
  var bg2 = style.getPropertyValue('--bg2').trim();

  var SHORT = {
    '国泰黄金ETF联接C': '黄金ETF联接C',
    '嘉实中证稀有金属主题ETF发起联接C': '稀有金属',
    '南方中证全指证券公司ETF联接C': '证券公司ETF',
    '华泰保兴安悦债券C': '债券（华泰保兴）',
    '易方达恒生港股通红利低波动ETF联接发起式C': '港股红利低波',
    '富国中证煤炭指数C': '煤炭指数',
    '招商中证白酒指数C': '白酒指数',
    '华泰柏瑞中证稀土产业ETF发起式联接C': '稀土产业',
    '嘉实上证科创板芯片ETF发起联接C': '科创芯片',
    '华宝中证沪港深新消费指数C': '新消费',
    '易方达中证1000量化增强C': '中证1000增强',
    '永赢医药创新智选混合发起C': '医药创新',
    '东方新能源汽车主题混合': '新能源汽车',
    '华夏中证动漫游戏ETF发起式联接C': '动漫游戏',
    '招商中证机器人ETF发起联接C': '机器人',
    '农银现代农业加混合': '现代农业',
    '富国中证军工指数(LOF)C': '军工指数',
    '天弘中证细分化工产业主题ETF联接C': '细分化工',
    '广发中证全指家用电器ETF联接C': '家用电器'
  };
  function short(n) { return SHORT[n] || n; }
  var axisCommon = {
    axisLine: { lineStyle: { color: rule } },
    axisTick: { show: false },
    axisLabel: { color: muted, fontSize: 11 }
  };

  // ---- 1. 权重对比 ----
  var idx = D.names.map(function (n, i) { return i; });
  var order = idx.slice().sort(function (a, b) { return (D.mvW[a] + D.rpW[a]) - (D.mvW[b] + D.rpW[b]); });
  var cats = order.map(function (i) { return short(D.names[i]); });
  var cw = echarts.init(document.getElementById('chart-weights'), null, { renderer: 'svg' });
  cw.setOption({
    animation: false,
    color: [accent, accent2],
    tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' }, appendToBody: true,
      valueFormatter: function (v) { return v.toFixed(2) + '%'; } },
    legend: { data: ['MV（最大夏普）', 'RP（ERC）'], textStyle: { color: ink }, top: 0 },
    grid: { left: 110, right: 40, top: 44, bottom: 20 },
    xAxis: { type: 'value', max: 45, axisLabel: { color: muted, formatter: '{value}%' },
      splitLine: { lineStyle: { color: rule } } },
    yAxis: { type: 'category', data: cats, axisLine: { lineStyle: { color: rule } },
      axisTick: { show: false }, axisLabel: { color: ink, fontSize: 11 } },
    series: [
      { name: 'MV（最大夏普）', type: 'bar', data: order.map(function (i) { return D.mvW[i]; }),
        barWidth: 7, itemStyle: { borderRadius: [0, 2, 2, 0] } },
      { name: 'RP（ERC）', type: 'bar', data: order.map(function (i) { return D.rpW[i]; }),
        barWidth: 7, itemStyle: { borderRadius: [0, 2, 2, 0] } }
    ]
  });
  window.addEventListener('resize', function () { cw.resize(); });

  // ---- 2. 今年以来累计收益 ----
  var cy = echarts.init(document.getElementById('chart-ytd'), null, { renderer: 'svg' });
  cy.setOption({
    animation: false,
    color: [accent, '#2F6F5F', accent2],
    tooltip: { trigger: 'axis', appendToBody: true,
      valueFormatter: function (v) { return v.toFixed(2) + '%'; } },
    legend: { data: ['MV', 'MV 带止损', 'RP'], textStyle: { color: ink }, top: 0 },
    grid: { left: 52, right: 24, top: 44, bottom: 56 },
    xAxis: { type: 'category', data: D.ytd.dates, axisLine: { lineStyle: { color: rule } },
      axisTick: { show: false }, axisLabel: { color: muted, fontSize: 10, interval: 24 } },
    yAxis: { type: 'value', axisLabel: { color: muted, formatter: '{value}%' },
      splitLine: { lineStyle: { color: rule } } },
    series: [
      { name: 'MV', type: 'line', data: D.ytd.mv, showSymbol: false, lineStyle: { width: 2 } },
      { name: 'MV 带止损', type: 'line', data: D.ytd.mvStop || [], showSymbol: false,
        lineStyle: { width: 2, type: 'dashed' } },
      { name: 'RP', type: 'line', data: D.ytd.rp, showSymbol: false, lineStyle: { width: 2 } }
    ]
  });
  window.addEventListener('resize', function () { cy.resize(); });

  // ---- 3. 近一年回撤 ----
  var cd = echarts.init(document.getElementById('chart-dd'), null, { renderer: 'svg' });
  cd.setOption({
    animation: false,
    color: [accent, accent2, '#2F6F5F'],
    tooltip: { trigger: 'axis', appendToBody: true,
      valueFormatter: function (v) { return v.toFixed(2) + '%'; } },
    legend: { data: ['MV', 'RP', 'MV 带止损'], textStyle: { color: ink }, top: 0 },
    grid: { left: 52, right: 24, top: 44, bottom: 56 },
    xAxis: { type: 'category', data: D.dd.dates, axisLine: { lineStyle: { color: rule } },
      axisTick: { show: false }, axisLabel: { color: muted, fontSize: 10, interval: 24 } },
    yAxis: { type: 'value', axisLabel: { color: muted, formatter: '{value}%' },
      splitLine: { lineStyle: { color: rule } } },
    series: [
      { name: 'MV', type: 'line', data: D.dd.mv, showSymbol: false,
        areaStyle: { opacity: 0.12 }, lineStyle: { width: 1.5 } },
      { name: 'RP', type: 'line', data: D.dd.rp, showSymbol: false,
        areaStyle: { opacity: 0.12 }, lineStyle: { width: 1.5 } },
      { name: 'MV 带止损', type: 'line', data: D.dd.mvStop || [], showSymbol: false,
        connectNulls: true, lineStyle: { width: 1.5, type: 'dashed' } }
    ]
  });
  window.addEventListener('resize', function () { cd.resize(); });

  // ---- 4. 止损方案累计收益 ----
  var cs = echarts.init(document.getElementById('chart-stop'), null, { renderer: 'svg' });
  cs.setOption({
    animation: false,
    color: [muted, accent2],
    tooltip: { trigger: 'axis', appendToBody: true,
      valueFormatter: function (v) { return v.toFixed(2) + '%'; } },
    legend: { data: ['无止损', '阶梯止损 10%/18%/25% + 冷静20日'], textStyle: { color: ink }, top: 0 },
    grid: { left: 56, right: 24, top: 44, bottom: 56 },
    xAxis: { type: 'category', data: D.stop.dates, axisLine: { lineStyle: { color: rule } },
      axisTick: { show: false }, axisLabel: { color: muted, fontSize: 10, interval: 24 } },
    yAxis: { type: 'value', axisLabel: { color: muted, formatter: '{value}%' },
      splitLine: { lineStyle: { color: rule } } },
    series: [
      { name: '无止损', type: 'line', data: D.stop.base, showSymbol: false, lineStyle: { width: 2 } },
      { name: '阶梯止损 10%/18%/25% + 冷静20日', type: 'line', data: D.stop.stop, showSymbol: false,
        lineStyle: { width: 2 } }
    ]
  });
  window.addEventListener('resize', function () { cs.resize(); });

  // ---- 5. 三区间表现 ----
  var cwin = echarts.init(document.getElementById('chart-windows'), null, { renderer: 'svg' });
  cwin.setOption({
    animation: false,
    color: [accent, accent2, '#2F6F5F'],
    tooltip: { trigger: 'axis', appendToBody: true,
      valueFormatter: function (v) { return v.toFixed(2) + '%'; } },
    legend: { data: ['MV', 'RP', 'MV 带止损'], textStyle: { color: ink }, top: 0 },
    grid: { left: 52, right: 24, top: 44, bottom: 32 },
    xAxis: { type: 'category', data: D.windows.labels, axisLine: { lineStyle: { color: rule } },
      axisTick: { show: false }, axisLabel: { color: ink } },
    yAxis: { type: 'value', axisLabel: { color: muted, formatter: '{value}%' },
      splitLine: { lineStyle: { color: rule } } },
    series: [
      { name: 'MV', type: 'bar', data: D.windows.mv, barWidth: 22,
        label: { show: true, position: 'bottom', color: muted, fontSize: 10,
          formatter: function (p) { return p.value.toFixed(2) + '%'; } } },
      { name: 'RP', type: 'bar', data: D.windows.rp, barWidth: 22,
        label: { show: true, position: 'bottom', color: muted, fontSize: 10,
          formatter: function (p) { return p.value.toFixed(2) + '%'; } } },
      { name: 'MV 带止损', type: 'bar', data: D.windows.mvStop || [], barWidth: 22,
        label: { show: true, position: 'bottom', color: muted, fontSize: 10,
          formatter: function (p) {
            return (p.value === null || p.value === undefined) ? '—' : p.value.toFixed(2) + '%';
          } } }
    ]
  });
  window.addEventListener('resize', function () { cwin.resize(); });

  // ---- 6. 相关系数矩阵（★ 模型选中 / ◇ 当前实际持有） ----
  if (D.corr && D.corr.names && D.corr.names.length) {
    var C = D.corr;
    var cLabels = C.names.map(function (n, i) {
      var tag = (C.sel[i] ? '★' : '') + (C.held[i] ? '◇' : '');
      return (tag ? tag + ' ' : '') + short(n);
    });
    var cData = [];
    C.matrix.forEach(function (row, y) {
      row.forEach(function (v, x) { if (v !== null) cData.push([x, y, v]); });
    });
    var ccorr = echarts.init(document.getElementById('chart-corr'), null, { renderer: 'svg' });
    ccorr.setOption({
      animation: false,
      tooltip: { appendToBody: true, formatter: function (p) {
        return C.names[p.data[1]] + ' × ' + C.names[p.data[0]] +
          '<br>' + C.method + ' 相关系数 <b>' + p.data[2].toFixed(2) + '</b>';
      } },
      grid: { left: 10, right: 84, top: 12, bottom: 10, containLabel: true },
      xAxis: { type: 'category', data: cLabels, axisLine: { lineStyle: { color: rule } },
        axisTick: { show: false },
        axisLabel: { color: ink, fontSize: 9, interval: 0, rotate: 45 } },
      yAxis: { type: 'category', data: cLabels, inverse: true, axisLine: { lineStyle: { color: rule } },
        axisTick: { show: false },
        axisLabel: { color: ink, fontSize: 9, interval: 0 } },
      visualMap: { min: -1, max: 1, calculable: false, orient: 'vertical',
        right: 8, top: 'middle', itemWidth: 14, itemHeight: 160,
        text: ['+1 正相关', '-1 负相关'], textStyle: { color: muted, fontSize: 10 },
        inRange: { color: ['#1E4F7A', '#4C86B0', '#F6F7F9', '#D98E7F', '#A33327'] } },
      series: [{ type: 'heatmap', data: cData,
        label: { show: true, fontSize: 8, color: '#FFFFFF',
          textBorderColor: 'rgba(0,0,0,0.55)', textBorderWidth: 1,
          formatter: function (p) { return p.data[2].toFixed(2); } },
        itemStyle: { borderColor: bg2 || '#FFFFFF', borderWidth: 1 },
        emphasis: { itemStyle: { borderColor: accent2, borderWidth: 2 } } }]
    });
    window.addEventListener('resize', function () { ccorr.resize(); });
  }

  // ---- 7. 真实持仓固定权重 vs 主要指数（YTD） ----
  if (D.ytdIdx && D.ytdIdx.dates && D.ytdIdx.dates.length) {
    var I = D.ytdIdx;
    var idxSeries = [
      { name: '组合（真实持仓固定权重）', key: 'port', color: accent2, width: 3, z: 6 },
      { name: '上证指数', key: 'sh', color: accent, width: 1.6, z: 3 },
      { name: '沪深300', key: 'hs300', color: '#4C86B0', width: 1.6, z: 3 },
      { name: '中证1000', key: 'zz1000', color: '#6E5A9E', width: 1.6, z: 3 },
      { name: '创业板指', key: 'cyb', color: '#3E8E7E', width: 1.6, z: 3 }
    ];
    var cit = echarts.init(document.getElementById('chart-ytd-idx'), null, { renderer: 'svg' });
    cit.setOption({
      animation: false,
      tooltip: { trigger: 'axis', appendToBody: true,
        valueFormatter: function (v) { return v.toFixed(2) + '%'; } },
      legend: { data: idxSeries.map(function (s) { return s.name; }), textStyle: { color: ink }, top: 0 },
      grid: { left: 56, right: 100, top: 48, bottom: 56 },
      xAxis: { type: 'category', data: I.dates, axisLine: { lineStyle: { color: rule } },
        axisTick: { show: false }, axisLabel: { color: muted, fontSize: 10, interval: 24 } },
      yAxis: { type: 'value', axisLabel: { color: muted, formatter: '{value}%' },
        splitLine: { lineStyle: { color: rule } } },
      series: idxSeries.map(function (s) {
        return { name: s.name, type: 'line', data: I[s.key], showSymbol: false, z: s.z,
          color: s.color, lineStyle: { width: s.width },
          endLabel: { show: true, color: s.color, fontSize: 10,
            fontWeight: s.key === 'port' ? 'bold' : 'normal',
            formatter: function (p) { return (p.value > 0 ? '+' : '') + p.value.toFixed(2) + '%'; } } };
      })
    });
    window.addEventListener('resize', function () { cit.resize(); });
  }
})();
