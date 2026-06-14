/* 服务达人挑战赛（移动端）—— 前端交互与模拟数据 */
(function () {
  'use strict';

  const MONTHS = ['2026-01', '2026-02', '2026-03', '2026-04', '2026-05', '2026-06'];
  const CURRENT_MONTH = MONTHS[MONTHS.length - 1];
  const PAGE_SIZE = 6;

  const NICK_PREFIX = ['快乐', '阳光', '微笑', '幸运', '安静', '勤奋', '温暖', '自由', '元气', '清风',
    '星辰', '海风', '麦田', '柠檬', '青柠', '木棉', '远山', '听雨', '向阳', '小鹿'];
  const NICK_SUFFIX = ['的猫', '小栈', '同学', '阿哲', '先生', '小姐', '日记', '物语', '不停', '一号',
    '宝贝', '酱', '大人', '本尊', 'er', '君', '酱紫', '森林', '时光', '可乐'];

  // 固定种子伪随机，保证每次刷新数据一致
  let seed = 20260614;
  function rand() { seed = (seed * 9301 + 49297) % 233280; return seed / 233280; }
  function randInt(min, max) { return Math.floor(rand() * (max - min + 1)) + min; }
  function pick(arr) { return arr[Math.floor(rand() * arr.length)]; }

  function generateMembers(count) {
    const list = [];
    for (let i = 0; i < count; i++) {
      const id = 'M2026' + String(1001 + i).padStart(4, '0');
      const bindMonthIdx = randInt(0, MONTHS.length - 1);
      const bindDay = randInt(1, 28);
      const bindDate = `${MONTHS[bindMonthIdx]}-${String(bindDay).padStart(2, '0')}`;
      const bindTime = `${bindDate} ${String(randInt(8, 22)).padStart(2, '0')}:${String(randInt(0, 59)).padStart(2, '0')}`;
      const points = {};
      MONTHS.forEach((m, idx) => {
        if (idx < bindMonthIdx) { points[m] = 0; return; }
        points[m] = rand() < 0.3 ? 0 : randInt(20, 680);
      });
      const total = Object.values(points).reduce((a, b) => a + b, 0);
      list.push({ id, nick: pick(NICK_PREFIX) + pick(NICK_SUFFIX), bindTime, bindDate, points, total });
    }
    return list;
  }

  const ALL_MEMBERS = generateMembers(46);

  const $ = (id) => document.getElementById(id);
  const els = {
    memberInput: $('memberInput'),
    monthChips: $('monthChips'),
    filterToggle: $('filterToggle'),
    filterDot: $('filterDot'),
    advPanel: $('advPanel'),
    bindStart: $('bindStart'),
    bindEnd: $('bindEnd'),
    sortSelect: $('sortSelect'),
    resetBtn: $('resetBtn'),
    applyBtn: $('applyBtn'),
    cardList: $('cardList'),
    emptyState: $('emptyState'),
    emptyResetBtn: $('emptyResetBtn'),
    loadMoreBtn: $('loadMoreBtn'),
    exportBtn: $('exportBtn'),
    listTitle: $('listTitle'),
    statMembers: $('statMembers'),
    statMonthPoints: $('statMonthPoints'),
    statTotalPoints: $('statTotalPoints'),
    statActive: $('statActive'),
  };

  const state = {
    member: '',
    month: CURRENT_MONTH,
    bindStart: '',
    bindEnd: '',
    sort: 'month-desc',
    visible: PAGE_SIZE,
  };

  function fmtNum(n) { return n.toLocaleString('zh-CN'); }

  // ---- 月份 chips ----
  function initMonthChips() {
    const items = [{ v: 'all', label: '全部累计' }]
      .concat(MONTHS.map((m) => ({ v: m, label: m.slice(5) + ' 月' })));
    els.monthChips.innerHTML = items.map((it) =>
      `<button class="chip ${it.v === state.month ? 'chip--active' : ''}" data-month="${it.v}">${it.label}</button>`
    ).join('');
    els.monthChips.querySelectorAll('.chip').forEach((c) => {
      c.addEventListener('click', () => {
        state.month = c.dataset.month;
        state.visible = PAGE_SIZE;
        els.monthChips.querySelectorAll('.chip').forEach((x) => x.classList.remove('chip--active'));
        c.classList.add('chip--active');
        render();
      });
    });
    // 让默认选中的月份滚动到可见位置（仅水平滚动，不影响页面）
    const active = els.monthChips.querySelector('.chip--active');
    if (active) els.monthChips.scrollLeft = active.offsetLeft - 12;
  }

  function getFiltered() {
    let rows = ALL_MEMBERS.slice();
    if (state.member.trim()) {
      const kw = state.member.trim().toLowerCase();
      rows = rows.filter((r) => r.id.toLowerCase().includes(kw) || r.nick.toLowerCase().includes(kw));
    }
    if (state.bindStart) rows = rows.filter((r) => r.bindDate >= state.bindStart);
    if (state.bindEnd) rows = rows.filter((r) => r.bindDate <= state.bindEnd);

    const monthVal = (r) => (state.month === 'all' ? r.total : (r.points[state.month] || 0));
    const sorters = {
      'month-desc': (a, b) => monthVal(b) - monthVal(a),
      'month-asc': (a, b) => monthVal(a) - monthVal(b),
      'total-desc': (a, b) => b.total - a.total,
      'total-asc': (a, b) => a.total - b.total,
      'bind-desc': (a, b) => b.bindTime.localeCompare(a.bindTime),
      'bind-asc': (a, b) => a.bindTime.localeCompare(b.bindTime),
    };
    rows.sort(sorters[state.sort] || sorters['month-desc']);
    return rows;
  }

  function renderStats(rows) {
    const monthPoints = rows.reduce((s, r) => s + (r.points[CURRENT_MONTH] || 0), 0);
    const totalPoints = rows.reduce((s, r) => s + r.total, 0);
    const activeCount = rows.filter((r) => (r.points[CURRENT_MONTH] || 0) > 0).length;
    els.statMembers.textContent = fmtNum(rows.length);
    els.statMonthPoints.textContent = fmtNum(monthPoints);
    els.statTotalPoints.textContent = fmtNum(totalPoints);
    els.statActive.textContent = fmtNum(activeCount);
  }

  function monthLabel() {
    return state.month === 'all' ? '累计积分' : (state.month.slice(5) + ' 月积分');
  }

  function render() {
    const rows = getFiltered();
    renderStats(rows);

    els.listTitle.textContent = `绑定会员明细 · 共 ${fmtNum(rows.length)} 人`;
    els.filterDot.hidden = !(state.bindStart || state.bindEnd || state.sort !== 'month-desc');

    if (rows.length === 0) {
      els.cardList.innerHTML = '';
      els.emptyState.hidden = false;
      els.loadMoreBtn.hidden = true;
      return;
    }
    els.emptyState.hidden = true;

    const shown = rows.slice(0, state.visible);
    els.cardList.innerHTML = shown.map((r, i) => {
      const rankNum = i + 1;
      const rankCls = rankNum <= 3 ? `rank-badge rank-badge--${rankNum}` : 'rank-badge';
      const monthPts = state.month === 'all' ? r.total : (r.points[state.month] || 0);
      const isActive = (r.points[CURRENT_MONTH] || 0) > 0;
      const hotCls = monthPts >= 400 ? ' hot' : '';
      return `
        <li class="mcard">
          <div class="mcard__top">
            <span class="${rankCls}">${rankNum}</span>
            <div class="mcard__id">
              <div class="mcard__id-num">${r.id}</div>
              <div class="mcard__nick">${r.nick}</div>
            </div>
            <span class="tag ${isActive ? 'tag--active' : 'tag--sleep'}">${isActive ? '本月活跃' : '本月沉睡'}</span>
          </div>
          <div class="mcard__bind"><span class="k">绑定时间</span> ${r.bindTime}</div>
          <div class="mcard__pts">
            <div class="pt-box">
              <div class="pt-box__label">${monthLabel()}</div>
              <div class="pt-box__value${hotCls}">${monthPts > 0 ? '+' + fmtNum(monthPts) : '—'}</div>
            </div>
            <div class="pt-box pt-box--total">
              <div class="pt-box__label">累计贡献积分</div>
              <div class="pt-box__value">${fmtNum(r.total)}</div>
            </div>
          </div>
        </li>`;
    }).join('');

    els.loadMoreBtn.hidden = state.visible >= rows.length;
    els.loadMoreBtn.textContent = `加载更多（剩余 ${fmtNum(Math.max(0, rows.length - state.visible))} 人）`;
  }

  function applyAdv() {
    state.member = els.memberInput.value;
    state.bindStart = els.bindStart.value;
    state.bindEnd = els.bindEnd.value;
    state.sort = els.sortSelect.value;
    state.visible = PAGE_SIZE;
    els.advPanel.hidden = true;
    render();
  }

  function resetFilters() {
    state.member = '';
    state.month = 'all';
    state.bindStart = '';
    state.bindEnd = '';
    state.sort = 'month-desc';
    state.visible = PAGE_SIZE;
    els.memberInput.value = '';
    els.bindStart.value = '';
    els.bindEnd.value = '';
    els.sortSelect.value = 'month-desc';
    els.monthChips.querySelectorAll('.chip').forEach((c) =>
      c.classList.toggle('chip--active', c.dataset.month === 'all'));
    els.advPanel.hidden = true;
    render();
  }

  function exportCsv() {
    const rows = getFiltered();
    const header = ['会员号', '会员昵称', '绑定时间', monthLabel(), '累计贡献积分'];
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

  els.filterToggle.addEventListener('click', () => { els.advPanel.hidden = !els.advPanel.hidden; });
  els.applyBtn.addEventListener('click', applyAdv);
  els.resetBtn.addEventListener('click', resetFilters);
  els.emptyResetBtn.addEventListener('click', resetFilters);
  els.exportBtn.addEventListener('click', exportCsv);
  els.loadMoreBtn.addEventListener('click', () => { state.visible += PAGE_SIZE; render(); });
  els.memberInput.addEventListener('input', () => { state.member = els.memberInput.value; state.visible = PAGE_SIZE; render(); });
  els.memberInput.addEventListener('keydown', (e) => { if (e.key === 'Enter') els.memberInput.blur(); });

  initMonthChips();
  render();
})();
