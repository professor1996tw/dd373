function doLogout() { location.reload(); }

// ===================================================
// 商品類型對照
// ===================================================
const ITEM_TYPE_MAP = {'游戏币':'遊戲幣','天幣':'天幣','点券':'點券','装备':'裝備','账号':'帳號'};

// ===================================================
// State
// ===================================================
let allOrders   = [];
let groupedData = [];
let fetchAborted = false;
let activeFetchTabId = null;

// ===================================================
// 訂單成功狀態判斷
// ===================================================
function isSuccess(status) {
  return ['交易成功','发货完成','准备发货'].includes(status);
}

// ===================================================
// 比例計算
// ===================================================
function getDiamondRate() {
  const b = parseFloat(document.getElementById('baseRate').value);
  const m = parseFloat(document.getElementById('rateMultiplier').value);
  return (isNaN(b) ? 4.692 : b) * (isNaN(m) ? 1.02 : m);
}
function updateRateCalc() {
  document.getElementById('calcRateDisplay').textContent = getDiamondRate().toFixed(3);
}

// 台幣成本：小數點後2位
function fmtTWD(n) {
  return Number(Math.round(n + 'e2') + 'e-2').toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}
// 進貨價：四捨五入整數
function fmtCost(n) {
  return Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

// ===================================================
// 時間範圍
// ===================================================
function toLocalDatetimeStr(d) {
  const p = n => String(n).padStart(2,'0');
  return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
}
function parseInputDatetime(str) {
  if (!str) return null;
  const d = new Date(str.trim().replace(' ','T'));
  return isNaN(d.getTime()) ? null : d;
}
function setTimeRange(r) {
  const now = new Date();
  let s, e;
  if (r==='today') {
    s=new Date(now.getFullYear(),now.getMonth(),now.getDate(),0,0,0);
    e=new Date(now.getFullYear(),now.getMonth(),now.getDate(),23,59,59);
  } else if (r==='yesterday') {
    const y=new Date(now); y.setDate(y.getDate()-1);
    s=new Date(y.getFullYear(),y.getMonth(),y.getDate(),0,0,0);
    e=new Date(y.getFullYear(),y.getMonth(),y.getDate(),23,59,59);
  } else if (r==='7d') {
    s=new Date(now); s.setDate(s.getDate()-6); s.setHours(0,0,0,0);
    e=new Date(now.getFullYear(),now.getMonth(),now.getDate(),23,59,59);
  } else {
    document.getElementById('startTime').value='';
    document.getElementById('endTime').value=''; return;
  }
  document.getElementById('startTime').value=toLocalDatetimeStr(s);
  document.getElementById('endTime').value=toLocalDatetimeStr(e);
}
function getTimeFilter() {
  const s = document.getElementById('startTime').value;
  const e = document.getElementById('endTime').value;
  return { start: s ? parseInputDatetime(s) : null, end: e ? parseInputDatetime(e) : null };
}
function parseOrderDate(str) {
  if (!str) return null;
  const d = new Date(str.replace(' ','T'));
  return isNaN(d.getTime()) ? null : d;
}
function applyTimeFilter(orders) {
  const {start,end} = getTimeFilter();
  if (!start && !end) return orders;
  return orders.filter(o => {
    const d = parseOrderDate(o.date);
    if (!d) return false;
    if (start && d < start) return false;
    if (end   && d > end)   return false;
    return true;
  });
}

// ===================================================
// 解析 goods
// ===================================================
function parseGoods(g) {
  if (!g) return null;
  const pm = g.match(/=(\d+\.?\d*)元/);
  const diamond = parseFloat(pm?.[1] || 0);
  const qm = g.match(/^(\d+\.?\d*)(萬|万)?/);
  let qty = 0;
  if (qm) qty = Math.round(parseFloat(qm[1]) * (qm[2] ? 10000 : 1));

  const im = g.match(/商品[类類]型[：:]\s*(.+?)(\s*$)/);
  const itemRaw = (im?.[1] || '').trim();
  const item = ITEM_TYPE_MAP[itemRaw] || itemRaw;

  const rm = g.match(/[游遊][戲戏][区區]服[：:]\s*(.+?)(?:\s+商品[类類]型|$)/);
  let parts = (rm?.[1] || '').trim().replace(/[：:]/g,'/').split('/').map(s=>s.trim()).filter(Boolean);

  if (parts.length === 3 && parts[0].includes(' ')) {
    const sub = parts[0].split(/\s+/);
    parts = [...sub, ...parts.slice(1)];
  }
  return { qty, item, diamond, game: parts.slice(0,2).join(''), serverName: parts[3] || '' };
}

// ===================================================
// 分組
// ===================================================
function groupOrders(orders) {
  const map = {};
  orders.forEach(o => {
    const p = parseGoods(o.goods);
    if (!p || !p.game) return;
    const k = `${p.game}||${p.item}||${p.serverName}`;
    if (!map[k]) map[k] = {game:p.game, item:p.item, serverName:p.serverName, qty:0, diamond:0, count:0};
    map[k].qty += p.qty; map[k].diamond += p.diamond; map[k].count++;
  });
  return Object.values(map).sort((a,b) => b.diamond - a.diamond);
}

// ===================================================
// 結算渲染
// ===================================================
function renderResults() {
  const diamondRate = getDiamondRate();
  const filtered = applyTimeFilter(allOrders);
  const successOrders = filtered.filter(o => isSuccess(o.status));

  const {start, end} = getTimeFilter();
  const infoEl = document.getElementById('timeRangeInfo');
  if (start || end) {
    const fmt = d => d ? d.toLocaleString('zh-TW') : '—';
    infoEl.textContent = `⏱ 時間範圍：${fmt(start)} ～ ${fmt(end)}　共 ${successOrders.length} 筆交易成功`;
    infoEl.style.display = 'block';
  } else { infoEl.style.display = 'none'; }

  const noDataEl = document.getElementById('noData');
  if (successOrders.length === 0) {
    noDataEl.textContent = '該條件下無交易成功訂單。請調整時間範圍或確認資料。';
    noDataEl.style.display = 'block';
    document.getElementById('resultContent').style.display = 'none';
    return;
  }

  noDataEl.style.display = 'none';
  document.getElementById('resultContent').style.display = 'block';

  const totalDia = successOrders.reduce((s,o) => {
    const m = o.goods?.match(/=(\d+\.?\d*)元/);
    return s + (m ? parseFloat(m[1]) : 0);
  }, 0);

  document.getElementById('sum-orders').textContent  = successOrders.length;
  document.getElementById('sum-diamond').textContent = totalDia.toFixed(2);
  document.getElementById('rateUsed').innerHTML = `鑽石比例 <span class="val">${diamondRate.toFixed(3)}</span>`;

  groupedData = groupOrders(successOrders);
  const tbody = document.getElementById('resultBody');
  tbody.innerHTML = '';
  let totQty=0, totDia=0, totTWD=0, totCost=0;

  groupedData.forEach(g => {
    const ratio = g.qty / g.diamond / 10000;
    const twd   = Number(Math.round(ratio * 10000 / diamondRate + 'e2') + 'e-2');
    const cost  = Number(Math.round(g.diamond * diamondRate + 'e2') + 'e-2');
    totQty += g.qty; totDia += g.diamond; totTWD += twd; totCost += cost;
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${g.game}</td><td>${g.item}</td><td>${g.serverName}</td>
      <td class="td-right td-mono">${g.qty.toLocaleString()}</td>
      <td class="td-right td-mono td-accent">${g.diamond.toFixed(2)}</td>
      <td class="td-right td-mono">${ratio.toFixed(4)}</td>
      <td class="td-right td-mono td-green">${fmtTWD(twd)}</td>
      <td class="td-right td-mono" style="color:var(--accent2);opacity:.75;">${fmtCost(cost)}</td>`;
    tbody.appendChild(tr);
  });

  const totRatio = totQty / totDia / 10000;
  const totRow = document.createElement('tr');
  totRow.className = 'total-row';
  totRow.innerHTML = `
    <td colspan="3">合計（${groupedData.length} 組 / ${successOrders.length} 筆）</td>
    <td class="td-right td-mono">${totQty.toLocaleString()}</td>
    <td class="td-right td-mono">${totDia.toFixed(2)}</td>
    <td class="td-right td-mono">${totRatio.toFixed(4)}</td>
    <td class="td-right td-mono">${fmtTWD(totTWD)}</td>
    <td class="td-right td-mono">${fmtCost(totCost)}</td>`;
  tbody.appendChild(totRow);

  document.getElementById('sum-twd').textContent = fmtTWD(totTWD);

  const rawBody = document.getElementById('rawBody');
  rawBody.innerHTML = '';
  successOrders.forEach(o => {
    const m = o.goods?.match(/=(\d+\.?\d*)元/);
    const dia = m ? parseFloat(m[1]) : 0;
    const goodsShort = (o.goods || '').split(/[游遊][戲戏][区區]服/)[0].trim();
    const tr2 = document.createElement('tr');
    tr2.innerHTML = `
      <td class="td-mono" style="font-size:12px;">${o.date||'—'}</td>
      <td class="td-mono" style="font-size:11px;">${o.orderNo||'—'}</td>
      <td>${goodsShort}</td>
      <td class="td-right td-mono td-accent">${dia.toFixed(2)}</td>
      <td><span class="pill pill-success">${o.status}</span></td>`;
    rawBody.appendChild(tr2);
  });
}

function reCalc() { if (allOrders.length > 0) renderResults(); }

// ===================================================
// Google 試算表匯出
// ===================================================
function exportGoogleSheets() {
  const dr = getDiamondRate();
  const headers = ['遊戲','品項','伺服器','數量','鑽石','比例','台幣成本','進貨價'];
  const rows = [headers];
  groupedData.forEach(g => {
    const ratio = g.qty / g.diamond / 10000;
    const twd   = Number(Math.round(ratio * 10000 / dr + 'e2') + 'e-2');
    const cost  = Number(Math.round(g.diamond * dr + 'e2') + 'e-2');
    rows.push([g.game, g.item, g.serverName, g.qty, g.diamond.toFixed(2), ratio.toFixed(4), twd.toFixed(2), Math.round(cost)]);
  });
  const tsv = rows.map(r => r.join('\t')).join('\n');
  navigator.clipboard.writeText(tsv).then(() => {
    window.open('https://sheets.new', '_blank');
    setTimeout(() => {
      alert('✅ 資料已複製到剪貼簿！\n\n在 Google 試算表中：\n1. 點選 A1\n2. Ctrl+V 貼上');
    }, 500);
  }).catch(() => {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob(['\uFEFF'+tsv], {type:'text/tab-separated-values;charset=utf-8;'}));
    a.download = `dd373-結算-${new Date().toISOString().slice(0,10)}.tsv`;
    a.click();
  });
}

// ===================================================
// 匯入到 admin
// ===================================================
function importToAdmin() {
  const dr = getDiamondRate();
  const payload = {
    data: groupedData.map(g => ({
      ...g,
      ratio: g.qty/g.diamond/10000,
      twd:  Number(Math.round(g.qty/g.diamond/10000*10000/dr+'e2')+'e-2'),
      cost: Number(Math.round(g.diamond*dr+'e2')+'e-2'),
      diamondRate: dr, source: 'DD373'
    })),
    rawOrders: allOrders.filter(o => isSuccess(o.status)),
    diamondRate: dr, importedAt: new Date().toISOString()
  };
  localStorage.setItem('sunwater_dd373_import', JSON.stringify(payload));
  alert(`✅ 已準備 ${payload.data.length} 筆分組進貨資料\n請開啟 sunwater-admin.html 並點選「讀取 DD373 匯入」。`);
}

// ===================================================
// 自動擷取 - Chrome Extension API
// ===================================================

// 這個函式會被注入到 DD373 頁面執行（必須完全自給自足）
function scrapeDd373Page() {
  var tables = [].slice.call(document.querySelectorAll('table')).filter(function(t) {
    return t.textContent.indexOf('DBA') > -1;
  });
  var orders = tables.map(function(t) {
    var rs = [].slice.call(t.querySelectorAll('tr'));
    var h = (rs[0] && rs[0].textContent || '').replace(/\s+/g, ' ').trim();
    var cells = [].slice.call(rs[1] && rs[1].querySelectorAll('td') || []).map(function(c) {
      return c.textContent.replace(/\s+/g, ' ').trim();
    });
    var statusRaw = cells[4] || '';
    var dateMatch  = h.match(/创建时间[：:]\s*(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})/);
    var orderMatch = h.match(/订单编号[：:]\s*(\w+)/);
    return {
      date:    dateMatch  ? dateMatch[1]  : '',
      orderNo: orderMatch ? orderMatch[1] : '',
      goods:   cells[0] || '',
      status:  statusRaw.replace(/\s+.*/, '')
    };
  }).filter(function(o) { return o.orderNo; });
  return orders;
}

function buildDd373Url(pageIndex, ddStartDate, pageSize) {
  const params = {
    TabStatus: '-1', Status: '-1',
    StartDate: ddStartDate,
    EndDate: '', Keyword: '', LastId: '', GoodsType: '',
    DealType: '-1', OrderBy: '1',
    PageIndex: String(pageIndex),
    PageSize:  String(pageSize || 80),
    RoleType: '1', IsRecycle: '0', timeId: '1'
  };
  return 'https://order.dd373.com/usercenter/buyer/buy_orders.html?searchParmList=' +
    encodeURIComponent(JSON.stringify(params));
}

function formatDateForDd373(date) {
  const p = n => String(n).padStart(2,'0');
  return `${date.getFullYear()}-${p(date.getMonth()+1)}-${p(date.getDate())} 00:00:00`;
}

function sleep(ms) { return new Promise(resolve => setTimeout(resolve, ms)); }

function waitForTabLoad(tabId) {
  return new Promise((resolve, reject) => {
    const timeout = setTimeout(() => {
      chrome.tabs.onUpdated.removeListener(listener);
      reject(new Error('頁面載入逾時（30秒）'));
    }, 30000);

    function listener(id, changeInfo) {
      if (id === tabId && changeInfo.status === 'complete') {
        clearTimeout(timeout);
        chrome.tabs.onUpdated.removeListener(listener);
        resolve();
      }
    }

    chrome.tabs.get(tabId).then(tab => {
      if (tab.status === 'complete') {
        clearTimeout(timeout);
        resolve();
      } else {
        chrome.tabs.onUpdated.addListener(listener);
      }
    }).catch(err => { clearTimeout(timeout); reject(err); });
  });
}

function setFetchUI(running) {
  document.getElementById('btnFetch').disabled = running;
  document.getElementById('btnAbort').style.display = running ? 'inline-block' : 'none';
}

function showFetchStatus(state, msg) {
  const el = document.getElementById('fetchStatus');
  el.className = 'fetch-status ' + (state === 'running' ? 'running' : state === 'done' ? 'done' : state === 'error' ? 'error' : '');
  const spin = state === 'running' ? '<span class="spinner"></span> ' : '';
  el.innerHTML = spin + msg;
}

async function startAutoFetch() {
  const startVal = document.getElementById('startTime').value;
  if (!startVal) {
    alert('請先設定起始時間');
    return;
  }

  // 鎖定結束時間
  const endEl = document.getElementById('endTime');
  if (!endEl.value) {
    endEl.value = toLocalDatetimeStr(new Date());
    endEl.style.borderColor = 'var(--accent)';
    setTimeout(() => { endEl.style.borderColor = ''; }, 3000);
  }

  const startTime = parseInputDatetime(startVal);
  if (!startTime) {
    alert('起始時間格式錯誤，請使用 2026-03-26 01:22:07 格式');
    return;
  }

  fetchAborted = false;
  setFetchUI(true);
  showFetchStatus('running', '正在開啟 DD373...');

  try {
    // DD373 起始日期 = 使用者起始時間前一天，確保不遺漏
    const ddStart = new Date(startTime.getTime() - 86400000);
    const ddStartStr = formatDateForDd373(ddStart);

    // 開啟 DD373 分頁（背景不聚焦）
    const tab = await chrome.tabs.create({ url: buildDd373Url(1, ddStartStr, 80), active: false });
    activeFetchTabId = tab.id;

    showFetchStatus('running', '等待頁面載入...');
    await waitForTabLoad(tab.id);

    // 強制刷新（相當於 Ctrl+F5）
    showFetchStatus('running', '🔄 強制刷新快取（Ctrl+F5）...');
    await chrome.tabs.reload(tab.id, { bypassCache: true });
    await waitForTabLoad(tab.id);
    await sleep(1000);

    // 逐頁爬取
    let scraped = [];
    let pageIndex = 1;
    let done = false;

    while (!done && !fetchAborted) {
      showFetchStatus('running', `📦 正在擷取第 ${pageIndex} 頁... （已累積 ${scraped.length} 筆）`);

      if (pageIndex > 1) {
        await chrome.tabs.update(tab.id, { url: buildDd373Url(pageIndex, ddStartStr, 80) });
        await waitForTabLoad(tab.id);
        await sleep(700);
      }

      let results;
      try {
        results = await chrome.scripting.executeScript({
          target: { tabId: tab.id },
          func: scrapeDd373Page
        });
      } catch (e) {
        showFetchStatus('error', '❌ 無法注入腳本，請確認已登入 DD373，並在 chrome://extensions 確認擴充功能已啟用');
        setFetchUI(false);
        return;
      }

      const pageOrders = (results && results[0] && results[0].result) || [];

      if (pageOrders.length === 0) {
        done = true;
      } else {
        const existingNos = new Set(scraped.map(o => o.orderNo));
        const newOrders = pageOrders.filter(o => !existingNos.has(o.orderNo));
        scraped = scraped.concat(newOrders);

        const oldest = pageOrders[pageOrders.length - 1];
        if (oldest && oldest.date) {
          const oldestDate = parseOrderDate(oldest.date);
          if (oldestDate && oldestDate < startTime) {
            done = true;
          }
        }

        if (pageOrders.length < 80) done = true;
        else if (!done) pageIndex++;
      }
    }

    if (!fetchAborted && activeFetchTabId) {
      await chrome.tabs.remove(activeFetchTabId).catch(() => {});
      activeFetchTabId = null;
    }

    if (!fetchAborted) {
      showFetchStatus('done', `✅ 擷取完成！共 ${scraped.length} 筆訂單，正在計算...`);
      allOrders = scraped;
      renderResults();
      showFetchStatus('done', `✅ 完成！共擷取 ${scraped.length} 筆，結果已更新`);
    }

  } catch (err) {
    showFetchStatus('error', `❌ 錯誤：${err.message}`);
  }

  setFetchUI(false);
}

function abortFetch() {
  fetchAborted = true;
  if (activeFetchTabId) {
    chrome.tabs.remove(activeFetchTabId).catch(() => {});
    activeFetchTabId = null;
  }
  showFetchStatus('', '已中止擷取');
  setFetchUI(false);
}

// ===================================================
// Init
// ===================================================
function initApp() {
  updateRateCalc();
  document.getElementById('btnToday').addEventListener('click', () => setTimeRange('today'));
  document.getElementById('btnYesterday').addEventListener('click', () => setTimeRange('yesterday'));
  document.getElementById('btn7d').addEventListener('click', () => setTimeRange('7d'));
  document.getElementById('btnClear').addEventListener('click', () => setTimeRange('all'));
  document.getElementById('baseRate').addEventListener('input', updateRateCalc);
  document.getElementById('rateMultiplier').addEventListener('input', updateRateCalc);
  document.getElementById('btnFetch').addEventListener('click', startAutoFetch);
  document.getElementById('btnAbort').addEventListener('click', abortFetch);
  document.getElementById('btnReCalc').addEventListener('click', reCalc);
  document.getElementById('btnExport').addEventListener('click', exportGoogleSheets);
  document.getElementById('btnImport').addEventListener('click', importToAdmin);
  document.getElementById('btnLogout').addEventListener('click', doLogout);
}

document.addEventListener('DOMContentLoaded', initApp);
