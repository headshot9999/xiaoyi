/* 服务达人挑战赛 —— 前端交互与模拟数据 */
(function () {
  'use strict';

  // ---- 赛季月份配置（最近 6 个月） ----
  const MONTHS = ['2026-01', '2026-02', '2026-03', '2026-04', '2026-05', '2026-06'];
  const CURRENT_MONTH = MONTHS[MONTHS.length - 1];
  const PAGE_SIZE = 8;

  const NICK_PREFIX = ['快乐', '阳光', '微笑', '幸运', '安静', '勤奋', '温暖', '自由', '元气', '清风',
    '星辰', '海风', '麦田', '柠檬', '青柠', '木棉', '远山', '听雨', '向阳', '小鹿'];
  const NICK_SUFFIX = ['的猫', '小栈', '同学', '阿哲', '先生', '小姐', '日记', '物语', '不停', '一号',
    '宝贝', '酱', '大人', '本尊', 'er', '君', '酱紫', '森林', '时光', '可乐'];

  // ---- 伪随机（固定种子，保证每次刷新数据一致，便于演示） ----
  let seed = 20260614;
  function rand() {
    seed = (seed * 9301 + 49297) % 233280;
    return seed / 233280;
  }
  function randInt(min, max) { return Math.floor(rand() * (max - min + 1)) + min; }
  function pick(arr) { return arr[Math.floor(rand() * arr.length)]; }

  // ---- 生成模拟会员数据 ----
  function generateMembers(count) {
    const list = [];
    for (let i = 0; i < count; i++) {
      const id = 'M2026' + String(1001 + i).padStart(4, '0');
      // 绑定时间：分布在赛季内
      const bindMonthIdx = randInt(0, MONTHS.length - 1);
      const bindDay = randInt(1, 28);
      const bindDate = `${MONTHS[bindMonthIdx]}-${String(bindDay).padStart(2, '0')}`;
      const bindTime = `${bindDate} ${String(randInt(8, 22)).padStart(2, '0')}:${String(randInt(0, 59)).padStart(2, '0')}`;

      // 每月贡献积分：绑定之后才会产生
      const points = {};
      MONTHS.forEach((m, idx) => {
        if (idx < bindMonthIdx) { points[m] = 0; return; }
        // 30% 概率某月不活跃
        points[m] = rand() < 0.3 ? 0 : randInt(20, 680);
      });
      const total = Object.values(points).reduce((a, b) => a + b, 0);

      list.push({
        id,
        nick: pick(NICK_PREFIX) + pick(NICK_SUFFIX),
        bindTime,
        bindDate,
        points,
        total,
      });
    }
    return list;
  }

  const ALL_MEMBERS = generateMembers(46);

  // ---- DOM ----
  const $ = (id) => document.getElementById(id);
  const els = {
    memberInput: $('memberInput'),
    monthSelect: $('monthSelect'),
    bindStart: $('bindStart'),
    bindEnd: $('bindEnd'),
    sortSelect: $('sortSelect'),
    searchBtn: $('searchBtn'),
    resetBtn: $('resetBtn'),
    activeFilters: $('activeFilters'),
    tableBody: $('tableBody'),
    emptyState: $('emptyState'),
    emptyResetBtn: $('emptyResetBtn'),
    resultCount: $('resultCount'),
    pagination: $('pagination'),
    monthColHead: $('monthColHead'),
    exportBtn: $('exportBtn'),
    statMembers: $('statMembers'),
    statMonthPoints: $('statMonthPoints'),
    statTotalPoints: $('statTotalPoints'),
    statActive: $('statActive'),
  };

  // ---- 状态 ----
  const state = {
    member: '',
    month: CURRENT_MONTH,
    bindStart: '',
    bindEnd: '',
    sort: 'month-desc',
    page: 1,
  };

  // ---- 初始化月份下拉 ----
  function initMonthOptions() {
    const opts = ['<option value="all">全部月份（累计）</option>']
      .concat(MONTHS.map((m) => `<option value="${m}">${m.replace('-', ' 年 ')} 月</option>`));
    els.monthSelect.innerHTML = opts.join('');
    els.monthSelect.value = state.month;
  }

  function fmtNum(n) { return n.toLocaleString('zh-CN'); }

  // ---- 筛选 + 排序 ----
  function getFiltered() {
    let rows = ALL_MEMBERS.slice();

    if (state.member.trim()) {
      const kw = state.member.trim().toLowerCase();
      rows = rows.filter((r) => r.id.toLowerCase().includes(kw) || r.nick.toLowerCase().includes(kw));
    }
    if (state.bindStart) rows = rows.filter((r) => r.bindDate >= state.bindStart);
    if (state.bindEnd) rows = rows.filter((r) => r.bindDate <= state.bindEnd);

    const monthVal = (r) => (state.month === 'all' ? r.total : (r.points[state.month] || 0));
    const bindVal = (r) => r.bindTime;

    const sorters = {
      'month-desc': (a, b) => monthVal(b) - monthVal(a),
      'month-asc': (a, b) => monthVal(a) - monthVal(b),
      'total-desc': (a, b) => b.total - a.total,
      'total-asc': (a, b) => a.total - b.total,
      'bind-desc': (a, b) => bindVal(b).localeCompare(bindVal(a)),
      'bind-asc': (a, b) => bindVal(a).localeCompare(bindVal(b)),
    };
    rows.sort(sorters[state.sort] || sorters['month-desc']);
    return rows;
  }

  // ---- 统计卡片（基于当前筛选结果） ----
  function renderStats(rows) {
    const monthPoints = rows.reduce((s, r) => s + (r.points[CURRENT_MONTH] || 0), 0);
    const totalPoints = rows.reduce((s, r) => s + r.total, 0);
    const activeCount = rows.filter((r) => (r.points[CURRENT_MONTH] || 0) > 0).length;
    els.statMembers.textContent = fmtNum(rows.length);
    els.statMonthPoints.textContent = fmtNum(monthPoints);
    els.statTotalPoints.textContent = fmtNum(totalPoints);
    els.statActive.textContent = fmtNum(activeCount);
  }

  // ---- 当前激活的筛选条件 chips ----
  function renderActiveChips() {
    const chips = [];
    if (state.member.trim()) chips.push({ key: 'member', label: `会员号：${state.member.trim()}` });
    if (state.month !== 'all') chips.push({ key: 'month', label: `月份：${state.month}` });
    else chips.push({ key: 'month', label: '月份：累计' });
    if (state.bindStart) chips.push({ key: 'bindStart', label: `绑定起：${state.bindStart}` });
    if (state.bindEnd) chips.push({ key: 'bindEnd', label: `绑定止：${state.bindEnd}` });

    els.activeFilters.innerHTML = chips.map((c) =>
      `<span class="chip">${c.label}<button data-key="${c.key}" aria-label="移除">×</button></span>`
    ).join('');

    els.activeFilters.querySelectorAll('button').forEach((btn) => {
      btn.addEventListener('click', () => {
        const key = btn.dataset.key;
        if (key === 'member') { state.member = ''; els.memberInput.value = ''; }
        if (key === 'month') { state.month = 'all'; els.monthSelect.value = 'all'; }
        if (key === 'bindStart') { state.bindStart = ''; els.bindStart.value = ''; }
        if (key === 'bindEnd') { state.bindEnd = ''; els.bindEnd.value = ''; }
        state.page = 1;
        render();
      });
    });
  }

  // ---- 表格渲染 ----
  function render() {
    const rows = getFiltered();
    renderStats(rows);
    renderActiveChips();

    els.monthColHead.textContent =
      state.month === 'all' ? '累计贡献积分' : `${state.month} 贡献积分`;
    els.resultCount.textContent = `共 ${fmtNum(rows.length)} 条`;

    const totalPages = Math.max(1, Math.ceil(rows.length / PAGE_SIZE));
    if (state.page > totalPages) state.page = totalPages;
    const start = (state.page - 1) * PAGE_SIZE;
    const pageRows = rows.slice(start, start + PAGE_SIZE);

    if (rows.length === 0) {
      els.tableBody.innerHTML = '';
      els.emptyState.hidden = false;
      els.pagination.innerHTML = '';
      return;
    }
    els.emptyState.hidden = true;

    els.tableBody.innerHTML = pageRows.map((r, i) => {
      const rankNum = start + i + 1;
      const rankCls = rankNum <= 3 ? `rank-badge rank-badge--${rankNum}` : 'rank-badge';
      const monthPts = state.month === 'all' ? r.total : (r.points[state.month] || 0);
      const isActive = (r.points[CURRENT_MONTH] || 0) > 0;
      const ptsCls = monthPts >= 400 ? 'points points--hot' : 'points';
      return `
        <tr>
          <td><span class="${rankCls}">${rankNum}</span></td>
          <td><span class="member-id">${r.id}</span></td>
          <td>${r.nick}</td>
          <td class="muted">${r.bindTime}</td>
          <td class="col-num"><span class="${ptsCls}">${monthPts > 0 ? '+' + fmtNum(monthPts) : '—'}</span></td>
          <td class="col-num"><span class="points-total">${fmtNum(r.total)}</span></td>
          <td class="col-status"><span class="tag ${isActive ? 'tag--active' : 'tag--sleep'}">${isActive ? '本月活跃' : '本月沉睡'}</span></td>
        </tr>`;
    }).join('');

    renderPagination(totalPages);
  }

  // ---- 分页控件 ----
  function renderPagination(totalPages) {
    if (totalPages <= 1) { els.pagination.innerHTML = ''; return; }
    const p = state.page;
    const parts = [];
    parts.push(`<button class="page-btn" data-page="${p - 1}" ${p === 1 ? 'disabled' : ''}>上一页</button>`);

    const pages = [];
    for (let i = 1; i <= totalPages; i++) {
      if (i === 1 || i === totalPages || (i >= p - 1 && i <= p + 1)) pages.push(i);
      else if (pages[pages.length - 1] !== '...') pages.push('...');
    }
    pages.forEach((i) => {
      if (i === '...') parts.push('<span class="page-ellipsis">…</span>');
      else parts.push(`<button class="page-btn ${i === p ? 'page-btn--active' : ''}" data-page="${i}">${i}</button>`);
    });

    parts.push(`<button class="page-btn" data-page="${p + 1}" ${p === totalPages ? 'disabled' : ''}>下一页</button>`);
    els.pagination.innerHTML = parts.join('');
    els.pagination.querySelectorAll('button[data-page]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const np = parseInt(btn.dataset.page, 10);
        if (!isNaN(np)) { state.page = np; render(); window.scrollTo({ top: 320, behavior: 'smooth' }); }
      });
    });
  }

  // ---- 事件绑定 ----
  function applyFiltersFromInputs() {
    state.member = els.memberInput.value;
    state.month = els.monthSelect.value;
    state.bindStart = els.bindStart.value;
    state.bindEnd = els.bindEnd.value;
    state.sort = els.sortSelect.value;
    state.page = 1;
    render();
  }

  function resetFilters() {
    state.member = '';
    state.month = 'all';
    state.bindStart = '';
    state.bindEnd = '';
    state.sort = 'month-desc';
    state.page = 1;
    els.memberInput.value = '';
    els.monthSelect.value = 'all';
    els.bindStart.value = '';
    els.bindEnd.value = '';
    els.sortSelect.value = 'month-desc';
    render();
  }

  function exportCsv() {
    const rows = getFiltered();
    const header = ['会员号', '会员昵称', '绑定时间',
      state.month === 'all' ? '累计贡献积分' : (state.month + ' 贡献积分'), '累计贡献积分'];
    const lines = [header.join(',')];
    rows.forEach((r) => {
      const monthPts = state.month === 'all' ? r.total : (r.points[state.month] || 0);
      lines.push([r.id, r.nick, r.bindTime, monthPts, r.total].join(','));
    });
    const blob = new Blob(['\ufeff' + lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `服务达人_绑定会员明细_${state.month}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  }

  els.searchBtn.addEventListener('click', applyFiltersFromInputs);
  els.resetBtn.addEventListener('click', resetFilters);
  els.emptyResetBtn.addEventListener('click', resetFilters);
  els.exportBtn.addEventListener('click', exportCsv);
  els.memberInput.addEventListener('input', applyFiltersFromInputs);
  els.memberInput.addEventListener('keydown', (e) => { if (e.key === 'Enter') applyFiltersFromInputs(); });
  els.monthSelect.addEventListener('change', applyFiltersFromInputs);
  els.sortSelect.addEventListener('change', applyFiltersFromInputs);
  els.bindStart.addEventListener('change', applyFiltersFromInputs);
  els.bindEnd.addEventListener('change', applyFiltersFromInputs);

  // ---- 启动 ----
  initMonthOptions();
  render();
})();
