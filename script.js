/* ─── Data ─────────────────────────────────────── */
const INITIAL_BALANCE = 4950;

let accounts = [
  {
    id: 1,
    name: 'Primary Manual',
    type: 'manual',
    startingBalance: INITIAL_BALANCE,
    trades: [
      {id:204529971, date:'2025-04-03', side:'sell', entry:3127.517, close:3127.826, sl:3133.949, rr:-0.05, pnl:-1.19,  lot:3.85,  bal:4948.81, notes:''},
      {id:204530950, date:'2025-04-03', side:'sell', entry:3125.095, close:3108.321, sl:3130.735, rr:2.97,  pnl:73.64,  lot:4.39,  bal:5022.45, notes:''},
      {id:204537982, date:'2025-04-04', side:'sell', entry:3097.119, close:3102.649, sl:3102.650, rr:-1.00, pnl:-24.72, lot:4.47,  bal:4997.73, notes:'SL hit'},
      {id:204539671, date:'2025-04-08', side:'buy',  entry:3004.324, close:2998.937, sl:2998.938, rr:-1.00, pnl:-24.78, lot:4.60,  bal:4972.95, notes:'SL hit'},
      {id:204540995, date:'2025-04-09', side:'buy',  entry:3018.845, close:3046.521, sl:3014.369, rr:6.18,  pnl:153.05, lot:5.53,  bal:5126.00, notes:'Best trade'},
      {id:204544149, date:'2025-04-11', side:'buy',  entry:3214.864, close:3215.805, sl:3207.941, rr:0.14,  pnl:3.37,   lot:3.58,  bal:5129.37, notes:''},
      {id:204547283, date:'2025-04-16', side:'buy',  entry:3292.905, close:3289.021, sl:3289.021, rr:-1.00, pnl:-49.48, lot:12.74, bal:5079.89, notes:'SL hit'},
      {id:204547881, date:'2025-04-17', side:'sell', entry:3333.170, close:3332.775, sl:3337.554, rr:0.09,  pnl:2.23,   lot:5.65,  bal:5082.12, notes:''},
      {id:204548837, date:'2025-04-21', side:'buy',  entry:3385.147, close:3384.845, sl:3381.359, rr:-0.08, pnl:-3.95,  lot:13.07, bal:5078.17, notes:''},
      {id:204549518, date:'2025-04-22', side:'sell', entry:3485.486, close:3452.204, sl:3500.001, rr:2.29,  pnl:113.49, lot:3.41,  bal:5191.66, notes:''},
      {id:204557213, date:'2025-04-30', side:'sell', entry:3307.337, close:3306.984, sl:3311.157, rr:0.09,  pnl:4.57,   lot:12.96, bal:5196.23, notes:''}
    ]
  },
  {
    id: 2,
    name: 'Excel Account',
    type: 'excel',
    startingBalance: INITIAL_BALANCE,
    trades: []
  }
];

let currentAccountId = 1;
const monthNames = ['January','February','March','April','May','June','July','August','September','October','November','December'];
let calendarState = (() => {
  const current = getCurrentTrades();
  const start = current.length ? new Date(current[0].date) : new Date();
  return { year: start.getFullYear(), month: start.getMonth() };
})();
let selectedCalendarDate = '';

function getCurrentAccount() {
  return accounts.find(a => a.id === currentAccountId) || accounts[0];
}

function getCurrentTrades() {
  return getCurrentAccount().trades;
}

function renderAccountSelector() {
  const account = getCurrentAccount();
  document.getElementById('accountControls').innerHTML = `
    <div style="display:flex;flex-wrap:wrap;gap:10px;align-items:center;">
      <label style="font-size:12px;color:var(--text-secondary);">Account</label>
      <select id="accountSelect" onchange="changeAccount(this.value)">
        ${accounts.map(a => `<option value="${a.id}" ${a.id === account.id ? 'selected' : ''}>${a.name} (${a.type})</option>`).join('')}
      </select>
      <button class="submit-btn" type="button" onclick="createAccount()">New account</button>
      <button class="secondary-btn" type="button" onclick="document.getElementById('excelFileInput').click()">Import Excel</button>
      <button class="danger-btn" type="button" onclick="deleteAccount()">Delete account</button>
    </div>
  `;
}

function changeAccount(id) {
  currentAccountId = parseInt(id, 10);
  selectedCalendarDate = '';
  render();
}

function createAccount() {
  const name = prompt('Account name', `Account ${accounts.length + 1}`);
  if (!name) return;
  const useExcel = confirm('Click OK to create an Excel-based account, or Cancel for a manual account.');
  const account = {
    id: Date.now(),
    name: name.trim() || `Account ${accounts.length + 1}`,
    type: useExcel ? 'excel' : 'manual',
    startingBalance: INITIAL_BALANCE,
    trades: []
  };
  accounts.push(account);
  currentAccountId = account.id;
  selectedCalendarDate = '';
  render();
}

function deleteAccount() {
  if (accounts.length === 1) {
    alert('You need at least one account.');
    return;
  }
  if (!confirm('Delete this account and all its trades?')) return;
  accounts = accounts.filter(a => a.id !== currentAccountId);
  currentAccountId = accounts[0].id;
  selectedCalendarDate = '';
  render();
}

function handleExcelUpload(event) {
  const file = event.target.files && event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    importTradesFromRows(rows);
    event.target.value = '';
  };
  reader.readAsArrayBuffer(file);
}

function normalizeDate(value) {
  if (!value) return '';
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return '';
  return parsed.toISOString().slice(0, 10);
}

function importTradesFromRows(rows) {
  if (!rows.length) {
    alert('No rows found in the imported file.');
    return;
  }
  const account = getCurrentAccount();
  const current = account.trades.slice();
  const startBalance = account.trades.length ? account.trades[account.trades.length - 1].bal : account.startingBalance;
  let runningBalance = startBalance;
  const imported = rows.map((row, index) => {
    const date = normalizeDate(row.Date || row.date || row['Trade Date'] || row['Date (YYYY-MM-DD)']);
    const side = String(row.Side || row.side || row.Action || '').toLowerCase();
    const entry = parseFloat(row.Entry || row.entry || row['Entry price'] || row['Entry Price'] || row['Price entry'] || 0);
    const close = parseFloat(row.Close || row.close || row['Exit'] || row['Close price'] || row['Close Price'] || 0);
    const sl = parseFloat(row.SL || row.sl || row['Stop Loss'] || row['Stop loss'] || 0) || 0;
    const pnl = parseFloat(row.Pnl || row.pnl || row['P&L'] || row['P/L'] || row['Realized P&L'] || row['Profit'] || 0);
    const lot = parseFloat(row.Lot || row.lot || row['Lot'] || row['Size'] || 0) || 0;
    const notes = String(row.Notes || row.notes || row['Comment'] || row['Memo'] || '').trim();
    if (!date || !side || isNaN(entry) || isNaN(close) || isNaN(pnl)) return null;
    runningBalance = parseFloat((runningBalance + pnl).toFixed(2));
    return {
      id: Date.now() + index,
      date,
      side: side === 'sell' ? 'sell' : 'buy',
      entry,
      close,
      sl,
      rr: sl > 0 ? parseFloat(((pnl >= 0 ? 1 : -1) * Math.abs(close - entry) / Math.abs(entry - sl)).toFixed(2)) : 0,
      pnl,
      lot,
      bal: runningBalance,
      notes
    };
  }).filter(Boolean);

  if (!imported.length) {
    alert('No valid trade rows were detected in the file. Use columns like Date, Side, Entry, Close, SL, Lot, and PnL.');
    return;
  }
  account.trades = current.concat(imported);
  accounts = accounts.map(a => a.id === account.id ? account : a);
  render();
}

function getCalendarMap() {
  return getCurrentTrades().reduce((map, trade) => {
    if (!trade.date) return map;
    const key = normalizeDate(trade.date);
    if (!key) return map;
    if (!map[key]) map[key] = [];
    map[key].push(trade);
    return map;
  }, {});
}

function renderCalendar() {
  const tradesByDate = getCalendarMap();
  const { year, month } = calendarState;
  const firstOfMonth = new Date(year, month, 1);
  const lastOfMonth = new Date(year, month + 1, 0);
  const dayNames = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  const startDay = firstOfMonth.getDay();
  const totalDays = lastOfMonth.getDate();

  let html = `<div class="calendar-header">
      <button class="secondary-btn" type="button" onclick="calendarState.month--; if (calendarState.month < 0) { calendarState.month = 11; calendarState.year--; } renderCalendar();">‹</button>
      <div>${monthNames[month]} ${year}</div>
      <button class="secondary-btn" type="button" onclick="calendarState.month++; if (calendarState.month > 11) { calendarState.month = 0; calendarState.year++; } renderCalendar();">›</button>
    </div>
    <div class="calendar-grid calendar-weekdays">
      ${dayNames.map(d => `<div>${d}</div>`).join('')}
    </div>
    <div class="calendar-grid calendar-days">
  `;

  for (let i = 0; i < startDay; i++) {
    html += '<div></div>';
  }

  for (let day = 1; day <= totalDays; day++) {
    const dateKey = `${year}-${String(month+1).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
    const dayTrades = tradesByDate[dateKey] || [];
    const isSelected = selectedCalendarDate === dateKey;
    const classes = ['calendar-day', dayTrades.length ? 'has-trades' : ''];
    if (isSelected) classes.push('selected');

    html += `<div class="${classes.filter(Boolean).join(' ')}" onclick="selectedCalendarDate='${dateKey}'; renderCalendar();">
      <div class="calendar-day-number">${day}</div>
      ${dayTrades.length ? `<div class="calendar-day-trades">${dayTrades.length} trade${dayTrades.length>1?'s':''}</div>` : ''}
    </div>`;
  }

  html += '</div>';

  if (selectedCalendarDate) {
    const selectedTrades = tradesByDate[selectedCalendarDate] || [];
    html += `<div class="calendar-details">
      <h3>Trades on ${selectedCalendarDate}</h3>
      ${selectedTrades.length ? selectedTrades.map(t => `
        <div class="calendar-trade-item">
          <strong>${t.side.toUpperCase()}</strong> ${t.entry} → ${t.close} (${t.pnl>=0?'+':''}$${t.pnl.toFixed(2)})
          <div style="font-size:12px;color:var(--text-secondary)">${t.notes || 'No note'}</div>
        </div>`).join('') : '<div>No trades for this day.</div>'}
    </div>`;
  }

  document.getElementById('calendarContainer').innerHTML = html;
}

/* ─── Metrics calculation ───────────────────────── */
function calcMetrics(arr) {
  const total = arr.reduce((s, t) => s + t.pnl, 0);
  const wins  = arr.filter(t => t.pnl > 0);
  const losses= arr.filter(t => t.pnl < 0);
  const wr    = arr.length ? (wins.length / arr.length * 100) : 0;
  const avgWin = wins.length  ? wins.reduce((s,t)=>s+t.pnl,0)/wins.length  : 0;
  const avgLoss= losses.length? losses.reduce((s,t)=>s+t.pnl,0)/losses.length: 0;
  const grossW = wins.reduce((s,t)=>s+t.pnl,0);
  const grossL = Math.abs(losses.reduce((s,t)=>s+t.pnl,0));
  const pf     = grossL > 0 ? grossW / grossL : grossW > 0 ? Infinity : 0;
  const avgRR  = arr.length ? arr.reduce((s,t)=>s+t.rr,0)/arr.length : 0;
  const finalBal = arr.length ? arr[arr.length-1].bal : INITIAL_BALANCE;
  const ret    = ((finalBal - INITIAL_BALANCE) / INITIAL_BALANCE * 100);
  const bestTrade = arr.length ? Math.max(...arr.map(t=>t.pnl)) : 0;
  const worstTrade= arr.length ? Math.min(...arr.map(t=>t.pnl)) : 0;
  return { total, wins:wins.length, losses:losses.length, wr, avgWin, avgLoss, pf, avgRR, finalBal, ret, bestTrade, worstTrade };
}

/* ─── Render metrics ────────────────────────────── */
function renderMetrics() {
  const trades = getCurrentTrades();
  const m = calcMetrics(trades);
  const items = [
    { label: 'Net P&L',       val: (m.total>=0?'+':'') + '$'+m.total.toFixed(2),   cls: m.total>=0?'pos':'neg' },
    { label: 'Win rate',      val: m.wr.toFixed(0)+'%',                             cls: m.wr>=50?'pos':'neg'  },
    { label: 'Trades',        val: trades.length,                                   cls: 'neu' },
    { label: 'Wins / Losses', val: m.wins+' / '+m.losses,                           cls: 'neu' },
    { label: 'Profit factor', val: isFinite(m.pf)?m.pf.toFixed(2):'∞',             cls: m.pf>=1?'pos':'neg'  },
    { label: 'Avg RR',        val: m.avgRR.toFixed(2)+'R',                          cls: m.avgRR>=1?'pos':'neg'},
    { label: 'Best trade',    val: '+$'+m.bestTrade.toFixed(2),                     cls: 'pos' },
    { label: 'Return',        val: (m.ret>=0?'+':'')+m.ret.toFixed(2)+'%',          cls: m.ret>=0?'pos':'neg'  },
  ];
  document.getElementById('metricsGrid').innerHTML = items.map(i =>
    `<div class="metric">
       <div class="metric-label">${i.label}</div>
       <div class="metric-val ${i.cls}">${i.val}</div>
     </div>`
  ).join('');
}

/* ─── Render table ──────────────────────────────── */
function renderTable() {
  const trades = getCurrentTrades();
  const side   = document.getElementById('filterSide').value;
  const result = document.getElementById('filterResult').value;
  let filtered = trades.filter(t => {
    if (side !== 'all' && t.side !== side) return false;
    if (result === 'win' && t.pnl <= 0) return false;
    if (result === 'loss' && t.pnl >= 0) return false;
    return true;
  });

  document.getElementById('tradeBody').innerHTML = filtered.map((t, i) => {
    const pc = t.pnl > 0 ? 'pos' : t.pnl < 0 ? 'neg' : 'neu';
    const bc = t.side === 'buy' ? 'badge-buy' : 'badge-sell';
    return `<tr>
      <td style="color:var(--text-tertiary)">${i+1}</td>
      <td>${t.date}</td>
      <td><span class="badge ${bc}">${t.side}</span></td>
      <td>${t.entry.toFixed(3)}</td>
      <td>${t.close.toFixed(3)}</td>
      <td>${t.sl.toFixed(3)}</td>
      <td style="color:var(--text-secondary)">${t.lot}</td>
      <td class="${pc}">${t.rr.toFixed(2)}R</td>
      <td class="${pc}">${t.pnl>=0?'+':''}$${t.pnl.toFixed(2)}</td>
      <td>$${t.bal.toFixed(2)}</td>
      <td class="note-cell" title="${t.notes}">${t.notes||'—'}</td>
      <td><button class="delete-btn" onclick="deleteTrade(${t.id})" title="Delete">✕</button></td>
    </tr>`;
  }).join('');
}

/* ─── Charts ────────────────────────────────────── */
let equityChart, pnlChart;

function renderCharts() {
  const trades = getCurrentTrades();
  const labels   = trades.map((_,i) => '#'+(i+1));
  const balances = [getCurrentAccount().startingBalance, ...trades.map(t => t.bal)];
  const pnls     = trades.map(t => t.pnl);
  const barColors= pnls.map(v => v >= 0 ? '#1D9E75' : '#D85A30');
  const isDark   = window.matchMedia('(prefers-color-scheme: dark)').matches;
  const gridColor= isDark ? 'rgba(255,255,255,0.07)' : 'rgba(0,0,0,0.07)';
  const tickColor= isDark ? '#6b6b67' : '#9b9b97';

  if (equityChart) equityChart.destroy();
  equityChart = new Chart(document.getElementById('equityChart'), {
    type: 'line',
    data: {
      labels: ['Start', ...labels],
      datasets: [{
        label: 'Balance',
        data: balances,
        borderColor: '#1D9E75',
        backgroundColor: 'rgba(29,158,117,0.08)',
        borderWidth: 1.5,
        pointRadius: 3,
        pointBackgroundColor: '#1D9E75',
        fill: true,
        tension: 0.3
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { font:{size:10}, color:tickColor, autoSkip:true, maxRotation:0 }, grid:{display:false} },
        y: { ticks: { font:{size:10}, color:tickColor, callback: v => '$'+v.toLocaleString() }, grid:{color:gridColor} }
      }
    }
  });

  if (pnlChart) pnlChart.destroy();
  pnlChart = new Chart(document.getElementById('pnlChart'), {
    type: 'bar',
    data: {
      labels,
      datasets: [{ label: 'P&L', data: pnls, backgroundColor: barColors, borderRadius: 3 }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { font:{size:10}, color:tickColor }, grid:{display:false} },
        y: { ticks: { font:{size:10}, color:tickColor, callback: v => (v>=0?'+':'')+' $'+v }, grid:{color:gridColor} }
      }
    }
  });
}

/* ─── Stats tab ─────────────────────────────────── */
function renderStats() {
  const trades = getCurrentTrades();
  const m = calcMetrics(trades);
  const buys  = trades.filter(t=>t.side==='buy');
  const sells = trades.filter(t=>t.side==='sell');
  const mBuy  = calcMetrics(buys);
  const mSell = calcMetrics(sells);

  document.getElementById('statsContent').innerHTML = `
    <table style="width:100%;border-collapse:collapse;">
      <tr style="border-bottom:0.5px solid var(--border)">
        <th style="text-align:left;padding:6px 10px;font-size:10px;color:var(--text-tertiary);text-transform:uppercase;letter-spacing:.05em">Metric</th>
        <th style="text-align:right;padding:6px 10px;font-size:10px;color:var(--text-tertiary);text-transform:uppercase;letter-spacing:.05em">Overall</th>
        <th style="text-align:right;padding:6px 10px;font-size:10px;color:var(--text-tertiary);text-transform:uppercase;letter-spacing:.05em">Buys</th>
        <th style="text-align:right;padding:6px 10px;font-size:10px;color:var(--text-tertiary);text-transform:uppercase;letter-spacing:.05em">Sells</th>
      </tr>
      ${[
        ['Total trades',   trades.length, buys.length, sells.length],
        ['Win rate',       m.wr.toFixed(0)+'%', mBuy.wr.toFixed(0)+'%', mSell.wr.toFixed(0)+'%'],
        ['Net P&L',        '$'+m.total.toFixed(2), '$'+mBuy.total.toFixed(2), '$'+mSell.total.toFixed(2)],
        ['Avg win',        '$'+m.avgWin.toFixed(2), '$'+mBuy.avgWin.toFixed(2), '$'+mSell.avgWin.toFixed(2)],
        ['Avg loss',       '$'+m.avgLoss.toFixed(2), '$'+mBuy.avgLoss.toFixed(2), '$'+mSell.avgLoss.toFixed(2)],
        ['Profit factor',  isFinite(m.pf)?m.pf.toFixed(2):'∞', isFinite(mBuy.pf)?mBuy.pf.toFixed(2):'∞', isFinite(mSell.pf)?mSell.pf.toFixed(2):'∞'],
        ['Avg RR',         m.avgRR.toFixed(2)+'R', mBuy.avgRR.toFixed(2)+'R', mSell.avgRR.toFixed(2)+'R'],
        ['Best trade',     '$'+m.bestTrade.toFixed(2), '$'+mBuy.bestTrade.toFixed(2), '$'+mSell.bestTrade.toFixed(2)],
        ['Worst trade',    '$'+m.worstTrade.toFixed(2), '$'+mBuy.worstTrade.toFixed(2), '$'+mSell.worstTrade.toFixed(2)],
        ['Return on acct', m.ret.toFixed(2)+'%', '—', '—'],
      ].map(row => `
        <tr style="border-bottom:0.5px solid var(--border)">
          <td style="padding:7px 10px;color:var(--text-secondary);font-size:12px">${row[0]}</td>
          <td style="padding:7px 10px;text-align:right;font-size:12px">${row[1]}</td>
          <td style="padding:7px 10px;text-align:right;font-size:12px;color:var(--green-text)">${row[2]}</td>
          <td style="padding:7px 10px;text-align:right;font-size:12px;color:var(--blue-text)">${row[3]}</td>
        </tr>`).join('')}
    </table>`;
}

/* ─── Tab switching ─────────────────────────────── */
function switchTab(tab) {
  ['log','add','stats','calendar'].forEach(id => {
    document.getElementById(id+'Tab').style.display = id === tab ? '' : 'none';
  });
  document.querySelectorAll('.tab').forEach((el, i) => {
    el.classList.toggle('active', ['log','add','stats','calendar'][i] === tab);
  });
  if (tab === 'stats') renderStats();
  if (tab === 'calendar') renderCalendar();
}

/* ─── Add trade ─────────────────────────────────── */
function addTrade() {
  const date  = document.getElementById('f-date').value;
  const side  = document.getElementById('f-side').value;
  const entry = parseFloat(document.getElementById('f-entry').value);
  const close = parseFloat(document.getElementById('f-close').value);
  const sl    = parseFloat(document.getElementById('f-sl').value) || 0;
  const pnl   = parseFloat(document.getElementById('f-pnl').value);
  const lot   = parseFloat(document.getElementById('f-lot').value) || 0;
  const notes = document.getElementById('f-notes').value.trim();

  if (!date || isNaN(entry) || isNaN(close) || isNaN(pnl)) {
    alert('Please fill in Date, Entry, Close, and P&L.');
    return;
  }

  const slDist   = sl > 0 ? Math.abs(entry - sl) : 0;
  const moveDist = Math.abs(close - entry);
  const rr       = slDist > 0 ? parseFloat(((pnl >= 0 ? 1 : -1) * moveDist / slDist).toFixed(2)) : 0;
  const current = getCurrentAccount();
  const prevBal  = current.trades.length ? current.trades[current.trades.length-1].bal : current.startingBalance;

  current.trades.push({
    id: Date.now(), date, side, entry, close, sl, rr, pnl, lot,
    bal: parseFloat((prevBal + pnl).toFixed(2)), notes
  });

  ['f-date','f-entry','f-close','f-sl','f-pnl','f-lot','f-notes'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });

  render();
  switchTab('log');
}

/* ─── Delete trade ──────────────────────────────── */
function deleteTrade(id) {
  if (!confirm('Delete this trade?')) return;
  const current = getCurrentAccount();
  current.trades = current.trades.filter(t => t.id !== id);
  let bal = current.startingBalance;
  current.trades.forEach(t => { bal = parseFloat((bal + t.pnl).toFixed(2)); t.bal = bal; });
  render();
}

/* ─── Main render ───────────────────────────────── */
function render() {
  renderAccountSelector();
  renderMetrics();
  renderTable();
  renderCharts();
  if (document.getElementById('statsTab').style.display !== 'none') renderStats();
  if (document.getElementById('calendarTab').style.display !== 'none') renderCalendar();
}

render();
