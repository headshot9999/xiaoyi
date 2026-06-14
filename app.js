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

  // 将某月总积分拆分为 k 笔订单（每笔 >= 1，且之和等于总积分）
  function splitPoints(total, k) {
    k = Math.max(1, Math.min(k, total));
    const parts = new Array(k).fill(1);
    let rem = total - k;
    while (rem > 0) {
      const i = Math.floor(rand() * k);
      const add = Math.min(rem, randInt(1, Math.max(1, Math.ceil(rem / 2))));
      parts[i] += add;
      rem -= add;
    }
    return parts;
  }

  function generateMembers(count) {
    const list = [];
    for (let i = 0; i < count; i++) {
      const id = 'M2026' + String(1001 + i).padStart(4, '0');
      const bindMonthIdx = randInt(0, MONTHS.length - 1);
      const bindDay = randInt(1, 28);
      const bindDate = `${MONTHS[bindMonthIdx]}-${String(bindDay).padStart(2, '0')}`;
      const bindTime = `${bindDate} ${String(randInt(8, 22)).padStart(2, '0')}:${String(randInt(0, 59)).padStart(2, '0')}`;
      const points = {};
      const orders = [];
      MONTHS.forEach((m, idx) => {
        if (idx < bindMonthIdx) { points[m] = 0; return; }
        const p = rand() < 0.3 ? 0 : randInt(20, 680);
        points[m] = p;
        if (p > 0) {
          const parts = splitPoints(p, randInt(1, 4));
          parts.forEach((pts) => {
            const day = randInt(1, 28);
            const hh = String(randInt(8, 22)).padStart(2, '0');
            const mm = String(randInt(0, 59)).padStart(2, '0');
            orders.push({
              no: 'AO' + m.replace('-', '') + String(randInt(10000, 99999)),
              datetime: `${m}-${String(day).padStart(2, '0')} ${hh}:${mm}`,
              month: m,
              points: pts,
            });
          });
        }
      });
      const total = Object.values(points).reduce((a, b) => a + b, 0);
      list.push({ id, nick: pick(NICK_PREFIX) + pick(NICK_SUFFIX), bindTime, bindDate, points, total, orders });
    }
    return list;
  }

  const MEMBERS = generateMembers(46);

  // 服务达人（本人）各月排名与总排名（模拟）
  const EXPERT_RANK = {
    months: { '2026-01': 12, '2026-02': 9, '2026-03': 7, '2026-04': 5, '2026-05': 8, '2026-06': 6 },
    total: 8,
  };

  const $ = (id) => document.getElementById(id);
  const els = {
    memberInput: $('memberInput'),
    monthPicker: $('monthPicker'),
    monthPickerWrap: $('monthPickerWrap'),
    monthAllBtn: $('monthAllBtn'),
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
    listTitle: $('listTitle'),
    statMembers: $('statMembers'),
    statMonthPoints: $('statMonthPoints'),
    statTotalPoints: $('statTotalPoints'),
    rankNum: $('rankNum'),
    rankLabel: $('rankLabel'),
    // 悬浮按钮 & 弹层
    bindFab: $('bindFab'),
    backdrop: $('backdrop'),
    detailSheet: $('detailSheet'),
    detailClose: $('detailClose'),
    detailId: $('detailId'),
    detailNick: $('detailNick'),
    detailBind: $('detailBind'),
    detailTotal: $('detailTotal'),
    detailMonths: $('detailMonths'),
    bindSheet: $('bindSheet'),
    bindClose: $('bindClose'),
    bindMemberId: $('bindMemberId'),
    bindMemberNick: $('bindMemberNick'),
    bindTip: $('bindTip'),
    bindConfirm: $('bindConfirm'),
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

  // ---- 月份筛选（日期控件 + 全部累计） ----
  function syncMonthControl() {
    const isAll = state.month === 'all';
    els.monthAllBtn.classList.toggle('is-active', isAll);
    els.monthPickerWrap.classList.toggle('is-disabled', isAll);
    if (!isAll) els.monthPicker.value = state.month;
  }
  function initMonthControl() {
    els.monthPicker.addEventListener('change', () => {
      const v = els.monthPicker.value;
      if (!v) return;
      state.month = v;
      state.visible = PAGE_SIZE;
      syncMonthControl();
      render();
    });
    // 点击日期框区域时，自动从「全部累计」切回具体月份
    els.monthPickerWrap.addEventListener('click', () => {
      if (state.month === 'all') {
        state.month = els.monthPicker.value || CURRENT_MONTH;
        state.visible = PAGE_SIZE;
        syncMonthControl();
        render();
      }
    });
    els.monthAllBtn.addEventListener('click', () => {
      state.month = 'all';
      state.visible = PAGE_SIZE;
      syncMonthControl();
      render();
    });
    syncMonthControl();
  }

  function getFiltered() {
    let rows = MEMBERS.slice();
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

  function renderStats() {
    const monthPoints = MEMBERS.reduce((s, r) => s + (r.points[CURRENT_MONTH] || 0), 0);
    const totalPoints = MEMBERS.reduce((s, r) => s + r.total, 0);
    els.statMembers.textContent = fmtNum(MEMBERS.length);
    els.statMonthPoints.textContent = fmtNum(monthPoints);
    els.statTotalPoints.textContent = fmtNum(totalPoints);
  }

  function monthLabel() {
    return state.month === 'all' ? '累计积分' : (state.month.slice(5) + ' 月积分');
  }

  // 右上角：当月排名 / 总排名（服务达人本人，紧凑展示）
  function renderExpertRank() {
    if (state.month === 'all') {
      els.rankNum.textContent = '第 ' + EXPERT_RANK.total + ' 名';
      els.rankLabel.textContent = '总排名';
    } else {
      const r = EXPERT_RANK.months[state.month];
      els.rankNum.textContent = '第 ' + (r || '—') + ' / ' + EXPERT_RANK.total + ' 名';
      els.rankLabel.textContent = state.month.slice(5) + ' 月 / 总排名';
    }
  }

  function render() {
    const rows = getFiltered();
    renderStats();
    renderExpertRank();

    els.listTitle.textContent = `绑定会员明细 · 共 ${fmtNum(rows.length)} 人`;
    els.filterDot.hidden = !(state.bindStart || state.bindEnd || state.sort !== 'month-desc');

    if (rows.length === 0) {
      els.cardList.innerHTML = '';
      els.emptyState.hidden = false;
      els.loadMoreBtn.hidden = true;
      return;
    }
    els.emptyState.hidden = true;
    const isAll = state.month === 'all';

    const shown = rows.slice(0, state.visible);
    els.cardList.innerHTML = shown.map((r) => {
      const monthPts = isAll ? r.total : (r.points[state.month] || 0);
      const hotCls = monthPts >= 400 ? ' hot' : '';
      return `
        <li class="mcard" data-id="${r.id}">
          <div class="mcard__top">
            <div class="mcard__id">
              <div class="mcard__id-num">${r.id}</div>
              <div class="mcard__nick">${r.nick}</div>
            </div>
            <div class="mcard__when">
              <span class="mcard__when-label">绑定时间</span>
              <span class="mcard__when-time">${r.bindTime}</span>
            </div>
            <span class="mcard__chev">›</span>
          </div>
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

    els.cardList.querySelectorAll('.mcard').forEach((card) => {
      card.addEventListener('click', () => openDetail(card.dataset.id));
    });

    els.loadMoreBtn.hidden = state.visible >= rows.length;
    els.loadMoreBtn.textContent = `加载更多（剩余 ${fmtNum(Math.max(0, rows.length - state.visible))} 人）`;
  }

  // ---- 弹层通用 ----
  function openSheet(sheet) {
    els.backdrop.hidden = false;
    sheet.hidden = false;
    sheet.setAttribute('aria-hidden', 'false');
    document.body.style.overflow = 'hidden';
  }
  function closeSheets() {
    els.backdrop.hidden = true;
    [els.detailSheet, els.bindSheet].forEach((s) => { s.hidden = true; s.setAttribute('aria-hidden', 'true'); });
    document.body.style.overflow = '';
  }

  // ---- 会员积分贡献明细 ----
  function openDetail(id) {
    const m = MEMBERS.find((x) => x.id === id);
    if (!m) return;
    els.detailId.textContent = m.id;
    els.detailNick.textContent = m.nick;
    els.detailBind.textContent = m.bindTime;
    els.detailTotal.textContent = fmtNum(m.total);

    const monthsWithOrders = MONTHS.slice().reverse()
      .filter((mo) => m.orders.some((o) => o.month === mo));

    if (monthsWithOrders.length === 0) {
      els.detailMonths.innerHTML = '<li class="detail-empty">该会员暂无积分贡献订单</li>';
    } else {
      els.detailMonths.innerHTML = monthsWithOrders.map((mo) => {
        const list = m.orders.filter((o) => o.month === mo)
          .sort((a, b) => b.datetime.localeCompare(a.datetime));
        const sum = list.reduce((s, o) => s + o.points, 0);
        const rows = list.map((o) => `
          <li class="order-item">
            <div class="order-item__main">
              <span class="order-item__no">订单号 ${o.no}</span>
              <span class="order-item__time">${o.datetime}</span>
            </div>
            <span class="order-item__pts">+${fmtNum(o.points)}</span>
          </li>`).join('');
        return `
          <li class="dm-group">
            <div class="dm-group__head">
              <span class="dm-group__month">${mo.replace('-', ' 年 ')} 月</span>
              <span class="dm-group__sum">合计 +${fmtNum(sum)}（${list.length} 笔）</span>
            </div>
            <ul class="order-list">${rows}</ul>
          </li>`;
      }).join('');
    }

    openSheet(els.detailSheet);
  }

  // ---- 绑定新会员 ----
  function openBind() {
    els.bindMemberId.value = '';
    els.bindMemberNick.value = '';
    els.bindTip.hidden = true;
    els.bindTip.className = 'bind-form__tip';
    openSheet(els.bindSheet);
    setTimeout(() => els.bindMemberId.focus(), 200);
  }

  function showTip(msg, type) {
    els.bindTip.textContent = msg;
    els.bindTip.className = 'bind-form__tip ' + type;
    els.bindTip.hidden = false;
  }

  function confirmBind() {
    const id = els.bindMemberId.value.trim();
    if (!id) { showTip('请输入会员号', 'err'); return; }
    if (MEMBERS.some((m) => m.id.toLowerCase() === id.toLowerCase())) {
      showTip('该会员号已绑定，请勿重复绑定', 'err');
      return;
    }
    const now = new Date();
    const pad = (n) => String(n).padStart(2, '0');
    const bindTime = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}`;
    const points = {};
    MONTHS.forEach((mo) => { points[mo] = 0; });
    MEMBERS.unshift({
      id,
      nick: els.bindMemberNick.value.trim() || '新会员',
      bindTime,
      bindDate: bindTime.slice(0, 10),
      points,
      total: 0,
      orders: [],
    });
    showTip('绑定成功！', 'ok');
    state.visible = PAGE_SIZE;
    render();
    setTimeout(closeSheets, 800);
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
    syncMonthControl();
    els.advPanel.hidden = true;
    render();
  }

  // ---- 事件绑定 ----
  els.filterToggle.addEventListener('click', () => { els.advPanel.hidden = !els.advPanel.hidden; });
  els.applyBtn.addEventListener('click', applyAdv);
  els.resetBtn.addEventListener('click', resetFilters);
  els.emptyResetBtn.addEventListener('click', resetFilters);
  els.loadMoreBtn.addEventListener('click', () => { state.visible += PAGE_SIZE; render(); });
  els.memberInput.addEventListener('input', () => { state.member = els.memberInput.value; state.visible = PAGE_SIZE; render(); });
  els.memberInput.addEventListener('keydown', (e) => { if (e.key === 'Enter') els.memberInput.blur(); });

  els.bindFab.addEventListener('click', openBind);
  els.bindConfirm.addEventListener('click', confirmBind);
  els.detailClose.addEventListener('click', closeSheets);
  els.bindClose.addEventListener('click', closeSheets);
  els.backdrop.addEventListener('click', closeSheets);

  initMonthControl();
  render();
})();
