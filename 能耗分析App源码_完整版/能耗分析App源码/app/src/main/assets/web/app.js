/* ================= 能耗分析系统 · 移动端 ================= */
const $ = s => document.querySelector(s);
const $$ = s => Array.from(document.querySelectorAll(s));
const S = { boot: null, page: 'dash', kind: 'elec', month: '', year: '', zones: [], dash: null, structBy: 'sys', catPick: [], rd: null, rdDirty: false };
const PALETTE = ['#1a56db', '#0ea5a4', '#f59e0b', '#8b5cf6', '#ec4899', '#16a34a', '#ef4444', '#0891b2', '#a16207', '#64748b'];
const NAV = [
  { k: 'dash', t: '总览', ic: '◎' }, { k: 'analysis', t: '分析', ic: '◈' }, { k: 'trend', t: '趋势', ic: '∿' },
  { k: 'meters', t: '表计', ic: '▤' }, { k: 'reading', t: '抄表', ic: '✎' }, { k: 'more', t: '更多', ic: '☰' },
];
const MORE = [
  { k: 'loss', t: '损耗分析', d: '国网 / 总表 / 分表三级比对' }, { k: 'quality', t: '抄表质量', d: '抄表间隔与异常画像' },
  { k: 'import', t: '数据导入', d: 'Excel 抄表 / 台账导入' }, { k: 'export', t: '报表导出', d: 'Excel 报表 / 数据备份' },
  { k: 'settings', t: '系统设置', d: '单价、分类、分组、密码' }, { k: 'help', t: '使用帮助', d: '口径说明与操作指引' },
];
const TITLE = {}; NAV.concat(MORE).forEach(n => TITLE[n.k] = n.t);

/* ---------- 工具 ---------- */
function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '<lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function num(v, d = 1) { if (v == null || v === '' || isNaN(v)) return '—'; return Number(v).toLocaleString('zh-CN', { minimumFractionDigits: d, maximumFractionDigits: d }); }
function money(v, d = 0) { if (v == null || isNaN(v)) return '—'; return Number(v).toLocaleString('zh-CN', { minimumFractionDigits: d, maximumFractionDigits: d }); }
function pct(v) { return v == null || isNaN(v) ? '—' : (v > 0 ? '+' : '') + Number(v).toFixed(1) + '%'; }
function pctCls(v) { if (v == null) return 'flat'; return v > 3 ? 'up' : (v < -3 ? 'down' : 'flat'); }
function pctTag(v) { return `<span class="tag ${pctCls(v)}">${pct(v)}</span>`; }
function unitOf(k) { return k === 'elec' ? 'kWh' : '吨'; }
function kindName(k) { return k === 'elec' ? '用电' : '用水'; }
function fmtMonth(mk) { return mk ? mk.slice(0, 4) + ' 年 ' + parseInt(mk.slice(5)) + ' 月' : ''; }
function toast(msg, type = '') { const t = $('#toast'); t.className = 'toast show ' + type; t.textContent = msg; clearTimeout(t._t); t._t = setTimeout(() => t.className = 'toast ' + type, 2400); }
function u_() { return unitOf(S.kind); }
function isE_() { return S.kind === 'elec'; }
function zoneQS() { return S.zones.length ? { zones: S.zones.join(',') } : {}; }

/* ---------- API ---------- */
async function api(path, data) {
  const opt = { headers: { 'Content-Type': 'application/json' }, credentials: 'same-origin' };
  if (data !== undefined) { opt.method = 'POST'; opt.body = JSON.stringify(data); }
  const r = await fetch('/api/' + path, opt);
  const j = await r.json().catch(() => ({ ok: false, msg: '返回格式错误' }));
  if (!j.ok) throw new Error(j.msg || '请求失败');
  return j.data;
}
const get = (p, q) => api(p + (q ? '?' + new URLSearchParams(q) : ''));
const post = (p, d) => api(p, d || {});

/* ---------- 弹层 ---------- */
function modal(title, body, foot) {
  const m = $('#modal');
  m.style.display = 'flex';
  m.innerHTML = `<div class="modal" onclick="event.stopPropagation()">
    <div class="modal-h"><h3>${esc(title)}</h3><span class="x" onclick="closeModal()">×</span></div>
    <div class="modal-b">${body}</div>
    <div class="modal-f">${foot || '<button class="btn" onclick="closeModal()">关闭</button>'}</div></div>`;
}
function closeModal() { $('#modal').style.display = 'none'; $('#modal').innerHTML = ''; }
$('#modal') && document.addEventListener('click', e => { if (e.target.id === 'modal') closeModal(); });
function sheet(title, body) {
  const m = $('#sheet');
  m.style.display = 'flex';
  m.innerHTML = `<div class="sheet" onclick="event.stopPropagation()">
    <div class="sheet-h"><b>${esc(title)}</b><span class="x" onclick="closeSheet()">×</span></div>
    <div class="sheet-b">${body}</div></div>`;
}
function closeSheet() { $('#sheet').style.display = 'none'; $('#sheet').innerHTML = ''; }
document.addEventListener('click', e => { if (e.target.id === 'sheet') closeSheet(); });
document.addEventListener('keydown', e => { if (e.key === 'Escape') { closeModal(); closeSheet(); } });

/* ---------- 图表（原生 SVG，触摸横滑） ---------- */
function niceNum(v) { const e = Math.floor(Math.log10(v || 1)), f = (v || 1) / Math.pow(10, e); let n = f <= 1 ? 1 : f <= 2 ? 2 : f <= 5 ? 5 : 10; return n * Math.pow(10, e); }
function shortNum(v) { if (v == null) return '—'; v = Math.abs(v); if (v >= 1e8) return (v / 1e8).toFixed(1) + '亿'; if (v >= 1e4) return (v / 1e4).toFixed(1) + '万'; if (v >= 1000) return (v / 1000).toFixed(1) + 'k'; return String(Math.round(v)); }
function drawChart(el, cfg) {
  if (!el) return;
  const labels = cfg.labels || [], series = (cfg.series || []).filter(s => s && s.data);
  const H = cfg.height || 220, n = labels.length;
  if (!n || !series.length) { el.innerHTML = '<div class="empty">暂无数据</div>'; return; }
  const minW = Math.max(300, n * (cfg.barWidth || 46) + 20);
  const W = Math.max(el.clientWidth || 360, minW);
  const pad = { l: 46, r: series.some(s => s.axis === 'r') ? 46 : 14, t: 26, b: 38 };
  const iw = W - pad.l - pad.r, ih = H - pad.t - pad.b;
  const L = series.filter(s => s.axis !== 'r'), R = series.filter(s => s.axis === 'r');
  const mx = a => Math.max(0, ...a.flatMap(s => s.data.map(v => v == null ? 0 : v))) || 1;
  const maxL = niceNum(mx(L)), maxR = niceNum(mx(R));
  const xw = iw / n, y0 = pad.t + ih;
  const bars = series.filter(s => s.type !== 'line'), lines = series.filter(s => s.type === 'line');
  let g = '';
  for (let i = 0; i <= 4; i++) {
    const y = y0 - ih * i / 4;
    g += `<line x1="${pad.l}" y1="${y}" x2="${W - pad.r}" y2="${y}" stroke="${i ? '#eef2f7' : '#cbd5e1'}"/>`;
    g += `<text x="${pad.l - 6}" y="${y + 4}" text-anchor="end" font-size="10" fill="#8497a8">${shortNum(maxL * i / 4)}</text>`;
    if (R.length) g += `<text x="${W - pad.r + 6}" y="${y + 4}" font-size="10" fill="#8497a8">${shortNum(maxR * i / 4)}</text>`;
  }
  const sc = (s, v) => { const m = s.axis === 'r' ? maxR : maxL; return y0 - Math.max(0, v) / m * ih; };
  labels.forEach((lb, i) => {
    const cx = pad.l + xw * i + xw / 2;
    g += `<text x="${cx}" y="${H - pad.b + 15}" text-anchor="middle" font-size="10" fill="#6b7f93">${esc(String(lb).length > 8 ? String(lb).slice(2) : lb)}</text>`;
  });
  const bw = Math.min(cfg.maxBar || 24, xw * 0.6 / Math.max(1, bars.length));
  bars.forEach((s, si) => {
    const off = (si - (bars.length - 1) / 2) * (bw + 2);
    s.data.forEach((v, i) => {
      if (v == null) return; const x = pad.l + xw * i + xw / 2 + off - bw / 2, y = sc(s, v);
      g += `<rect x="${x}" y="${y}" width="${bw}" height="${Math.max(1, y0 - y)}" rx="2" fill="${s.color}"><title>${esc(labels[i])} ${esc(s.name)}：${num(v)}</title></rect>`;
    });
  });
  lines.forEach(s => {
    let d = ''; s.data.forEach((v, i) => { if (v == null) return; d += (d ? 'L' : 'M') + (pad.l + xw * i + xw / 2) + ',' + sc(s, v); });
    g += `<path d="${d}" fill="none" stroke="${s.color}" stroke-width="2"/>`;
    s.data.forEach((v, i) => { if (v == null) return; g += `<circle cx="${pad.l + xw * i + xw / 2}" cy="${sc(s, v)}" r="3" fill="#fff" stroke="${s.color}" stroke-width="2"><title>${esc(labels[i])} ${esc(s.name)}：${num(v)}</title></circle>`; });
  });
  let lx = pad.l;
  series.forEach(s => {
    g += `<rect x="${lx}" y="7" width="9" height="9" rx="2" fill="${s.color}"/><text x="${lx + 12}" y="15" font-size="10.5" fill="#5b7189">${esc(s.name)}</text>`;
    lx += 24 + String(s.name).length * 10;
  });
  el.innerHTML = `<svg width="${W}" height="${H}" style="display:block">${g}</svg>`;
}
const C_ELEC = '#1a56db', C_WATER = '#0ea5a4';
function ckColor(c) { return c && /^#[0-9a-fA-F]{6}$/.test(c) ? c : C_ELEC; }
function donut(items) {
  const total = items.reduce((a, b) => a + b.usage, 0) || 1;
  const R = 58, r = 36, cx = 66, cy = 66; let a0 = -Math.PI / 2, p = '';
  items.slice(0, 8).forEach((it, i) => {
    const ang = it.usage / total * Math.PI * 2; if (ang <= 0) return; const a1 = a0 + ang;
    const P = (rr, a) => `${(cx + rr * Math.cos(a)).toFixed(1)},${(cy + rr * Math.sin(a)).toFixed(1)}`;
    p += `<path d="M${P(R, a0)} A${R},${R} 0 ${ang > Math.PI ? 1 : 0},1 ${P(R, a1)} L${P(r, a1)} A${r},${r} 0 ${ang > Math.PI ? 1 : 0},0 ${P(r, a0)} Z" fill="${PALETTE[i % PALETTE.length]}"><title>${esc(it.label)} ${num(it.usage)}</title></path>`;
    a0 = a1;
  });
  return `<svg width="132" height="132" style="flex:0 0 132px">${p}
    <text x="${cx}" y="${cy - 2}" text-anchor="middle" font-size="10" fill="#8497a8">合计</text>
    <text x="${cx}" y="${cy + 14}" text-anchor="middle" font-size="14" font-weight="700">${shortNum(total)}</text></svg>`;
}
function barList(items, opt = {}) {
  if (!items || !items.length) return '<div class="empty">暂无数据</div>';
  const max = Math.max(...items.map(i => Math.abs(i.usage)), 1);
  return items.map((i, idx) => {
    const w = Math.max(2, Math.abs(i.usage) / max * 100);
    return `<div class="bar-row"><div class="nm" title="${esc(i.label)}">${esc(i.label)}</div>
      <div class="track"><div class="fill" style="width:${w}%;background:${opt.multi ? PALETTE[idx % PALETTE.length] : (opt.color || C_ELEC)}"></div></div>
      <div class="vl">${num(i.usage)}${opt.unit || ''} ${i.mom != null ? `<span class="tag ${pctCls(i.mom)}" style="font-size:10px">${pct(i.mom)}</span>` : ''}</div></div>`;
  }).join('');
}

/* ---------- 顶部栏 / 导航 ---------- */
function shell() {
  const b = S.boot || {};
  const years = (b.years && b.years.length) ? b.years.slice() : [String(new Date().getFullYear())];
  if (S.year && !years.includes(S.year)) years.push(S.year);
  years.sort();
  const zones = b.zones || [];
  $('#app').innerHTML = `
  <div class="appbar">
    <div class="row1">
      <div style="flex:1;min-width:0"><div class="ttl">${esc((b.meta && b.meta.title) || '能耗分析系统')}</div>
      <div class="prj">${esc((b.meta && b.meta.project) || '')} · ${esc(b.version || '')}</div></div>
      <button class="iconbtn" onclick="logout()" title="退出">⏻</button>
    </div>
    <div class="period">
      <div class="seg"><button class="${isE_() ? 'on' : ''}" onclick="setKind('elec')">⚡ 用电</button><button class="${!isE_() ? 'on' : ''}" onclick="setKind('water')">💧 用水</button></div>
      <select onchange="setYear(this.value)">${years.map(y => `<option ${y === S.year ? 'selected' : ''}>${y}</option>`).join('')}</select>
      <select onchange="setMonth(this.value)">${Array.from({ length: 12 }, (_, i) => { const mk = S.year + '-' + String(i + 1).padStart(2, '0'); return `<option value="${mk}" ${mk === S.month ? 'selected' : ''}>${i + 1} 月</option>`; }).join('')}</select>
    </div>
    ${zones.length ? `<div class="chips"><span class="chip ${S.zones.length ? '' : 'on'}" onclick="clearZones()">全部区域</span>
      ${zones.map(z => `<span class="chip ${S.zones.includes(z) ? 'on' : ''}" onclick="toggleZone('${esc(z)}')">${esc(z)}</span>`).join('')}</div>` : ''}
  </div>
  <div class="wrap" id="page"><div class="sk" style="height:120px"></div><div class="sk" style="height:220px;margin-top:10px"></div></div>
  <nav class="tabbar">${NAV.map(x => `<a class="${(S.page === x.k || (x.k === 'more' && MORE.some(m => m.k === S.page))) ? 'on' : ''}" onclick="go('${x.k}')"><span class="ic">${x.ic}</span>${x.t}</a>`).join('')}</nav>`;
}
function setMonth(mk) { if (!mk) return; S.month = mk; if (mk.slice(0, 4) !== S.year) S.year = mk.slice(0, 4); render(); }
function setYear(y) { S.year = y; S.month = y + '-' + (S.month || '').slice(5); if (S.month.length !== 7) S.month = y + '-12'; render(); }
function toggleZone(z) { const i = S.zones.indexOf(z); i >= 0 ? S.zones.splice(i, 1) : S.zones.push(z); shell(); render(); }
function clearZones() { S.zones = []; shell(); render(); }
function setKind(k) { if (S.kind !== k) { S.kind = k; if (S.page === 'dash' && S.dash) S.dash = null; shell(); render(); } }
function go(p) {
  if (p === 'more') return openMore();
  if (confirmLeave()) return;
  S.page = p; shell(); render();
  window.scrollTo(0, 0);
}
function openMore() {
  sheet('更多功能', MORE.map(m => `<div class="detail-list" onclick="closeSheet();go('${m.k}')"><div class="it"><div class="t"><span>${m.t}</span><span style="color:#8497a8">›</span></div><div class="d">${m.d}</div></div></div>`).join(''));
}
function confirmLeave() {
  if (S.rdDirty && S.page === 'reading') { if (!confirm('当前抄表尚未保存，确定离开？')) return true; S.rdDirty = false; }
  return false;
}
function render() {
  const el = $('#page'); if (!el) return;
  const fn = ({ dash: viewDash, analysis: viewAnalysis, trend: viewTrend, loss: viewLoss, meters: viewMeters, reading: viewReading, import: viewImport, export: viewExport, quality: viewQuality, settings: viewSettings, help: viewHelp })[S.page];
  el.innerHTML = '<div class="sk" style="height:100px"></div><div class="sk" style="height:200px;margin-top:10px"></div>';
  Promise.resolve().then(() => fn(el)).catch(e => { el.innerHTML = `<div class="notice bad">加载出错：${esc(e.message)}</div>`; });
}

/* ---------- 登录 ---------- */
async function boot() {
  try {
    const st = await get('auth_state');
    if (st.need_setup) return loginScreen(true);
    if (!st.logged) return loginScreen(false);
    await loadAll();
  } catch (e) {
    $('#app').innerHTML = `<div class="boot">加载失败：${esc(e.message)}<br><br><button class="btn pri" onclick="location.reload()">重试</button></div>`;
  }
}
function loginScreen(setup) {
  $('#app').innerHTML = `<div class="login-wrap"><div class="login-box">
    <h1>${setup ? '欢迎使用' : '能耗分析系统'}</h1>
    <p>${setup ? '首次使用，请设置管理密码' : '请输入管理密码'}</p>
    ${setup ? `<div class="field"><label>设置密码（至少 4 位）</label><input class="inp" type="password" id="pwd1"></div>
      <div class="field"><label>再次输入</label><input class="inp" type="password" id="pwd2" onkeydown="if(event.key==='Enter')doSetup()"></div>` :
      `<div class="field"><label>密码</label><input class="inp" type="password" id="pwd1" onkeydown="if(event.key==='Enter')doLogin()"></div>`}
    <button class="btn pri" onclick="${setup ? 'doSetup()' : 'doLogin()'}">${setup ? '设置并进入' : '登录'}</button>
    <div style="text-align:center;font-size:11px;color:#8497a8;margin-top:14px">© 开心团团</div></div></div>`;
  setTimeout(() => { const e = $('#pwd1'); if (e) e.focus(); }, 200);
}
async function doSetup() {
  const a = $('#pwd1').value, b = $('#pwd2').value;
  if (a.length < 4) return toast('密码至少 4 位', 'err');
  if (a !== b) return toast('两次输入不一致', 'err');
  try { await post('auth_setup', { password: a }); toast('设置成功', 'ok'); await loadAll(); } catch (e) { toast(e.message, 'err'); }
}
async function doLogin() { try { await post('auth_login', { password: $('#pwd1').value }); await loadAll(); } catch (e) { toast(e.message, 'err'); } }
async function logout() { if (!confirm('确定退出？')) return; try { await post('auth_logout'); } catch (e) {} location.reload(); }
async function loadAll() {
  S.boot = await get('bootstrap');
  const d = await get('dashboard', { kind: S.kind });
  S.dash = d; S.month = d.month; S.year = (d.month || '').slice(0, 4) || String(new Date().getFullYear());
  shell(); render();
}

/* ================= 1. 总览 ================= */
async function viewDash(el) {
  if (!S.dash) S.dash = await get('dashboard', { kind: S.kind, month: S.month, ...zoneQS() });
  const d = S.dash, c = d.cmp || {}, u = u_(), st = d.sub_total || {}, cov = d.cover || {};
  const groups = (d.groups || []).filter(g => (g.months || []).some(m => m.usage));
  const rc = d.reclaimed;
  el.innerHTML = `
  <div class="kpis">
    <div class="kpi ${isE_() ? 'e' : 'w'}"><div class="lab">本月${kindName(S.kind)}量<br>${isE_() ? '低压进线柜总表' : '自来水总表'}</div>
      <div class="val">${num(c.cur)}<small>${u}</small></div>
      <div class="sub">环比 ${pctTag(c.mom)} 同比 ${c.yoy == null ? '<span class="tag flat">无同期</span>' : pctTag(c.yoy)}</div></div>
    <div class="kpi"><div class="lab">本月费用</div><div class="val">¥${money(c.cost)}</div>
      <div class="sub">日均 ${num(c.daily_cur, 2)} ${u} ${pctTag(c.mom_daily)}</div></div>
    <div class="kpi"><div class="lab">分表合计（对比）</div><div class="val">${num(st.sub)}<small>${u}</small></div>
      <div class="sub">${st.gap == null ? '对比总表' : `总表${st.gap > 0 ? '多' : '少'} ${Math.abs(st.gap).toFixed(1)}%`}</div></div>
    <div class="kpi"><div class="lab">抄表覆盖</div><div class="val">${num(cov.pct, 0)}<small>%</small></div>
      <div class="sub">${cov.covered}/${cov.total} 块表有数据</div></div>
  </div>
  <div class="card"><div class="card-h"><h3>${S.year} 年每月${kindName(S.kind)}</h3><span class="hint">柱=用量 线=日均×30</span><span class="sp"></span><a class="btn sm" onclick="go('trend')">完整趋势</a></div>
    <div class="card-b"><div id="chYear" class="chart tall"></div></div></div>
  <div class="card"><div class="card-h"><h3>${kindName(S.kind)}结构</h3>
    <select class="inp sm" style="width:auto" onchange="setStructBy(this.value)">
      <option value="sys" ${S.structBy !== 'cat' ? 'selected' : ''}>按系统</option>
      <option value="cat" ${S.structBy === 'cat' ? 'selected' : ''}>按表计</option></select></div>
    <div class="card-b" id="structBox">${structBox(d)}</div></div>
  <div class="card"><div class="card-h"><h3>各系统排名</h3><span class="hint">含环比</span></div>
    <div class="card-b">${barList(d.struct, { unit: ' ' + u, multi: true })}</div></div>
  <div class="card"><div class="card-h"><h3>TOP 表计</h3><span class="hint">本月耗量最高</span></div>
    <div class="card-b">${barList((d.top || []).filter(x => !x.is_total).slice(0, 10).map(x => ({ label: x.name, usage: x.usage, mom: x.mom })), { unit: ' ' + u })}</div></div>
  ${groups.length ? `<div class="card"><div class="card-h"><h3>分区核算对比</h3><span class="sp"></span><a class="btn sm" onclick="go('settings')">设置分组</a></div>
    <div class="card-b"><div id="chGroup" class="chart mid"></div>
    <div class="tbl-wrap"><table><thead><tr><th>分组</th>${groups[0].months.slice(-6).map(m => `<th class="ta-r">${parseInt(m.month.slice(5))}月</th>`).join('')}<th class="ta-r">近6月</th></tr></thead><tbody>
    ${groups.map(g => { const ms = g.months.slice(-6); const p = ms[ms.length - 2] || {}; const l = ms[ms.length - 1] || {};
      return `<tr><td><i style="display:inline-block;width:8px;height:8px;border-radius:2px;background:${ckColor(g.color)};margin-right:5px"></i>${esc(g.name)}</td>
      ${ms.map(m => `<td class="ta-r">${num(m.usage)}</td>`).join('')}<td class="ta-r"><b>${num(ms.reduce((a, b) => a + b.usage, 0))}</b> ${pctTag(p.usage ? (l.usage - p.usage) / p.usage * 100 : null)}</td></tr>`; }).join('')}
    </tbody></table></div></div></div>` : ''}
  ${rc && (rc.supply || rc.reclaimed) ? `<div class="card"><div class="card-h"><h3>中水核算</h3><span class="hint">近12个月 · 补水量按各区占比分摊</span></div>
    <div class="card-b"><div class="mini"><div><div class="k">中水补水（自来水）</div><div class="v" style="color:#1565c0">${num(rc.supply)}</div></div>
    <div><div class="k">各区中水用量</div><div class="v" style="color:#15803d">${num(rc.reclaimed)}</div></div>
    <div><div class="k">中水损耗</div><div class="v" style="color:#b45309">${num(rc.loss)}</div></div></div>
    ${(rc.zones || []).length ? `<div class="tbl-wrap" style="margin-top:8px"><table><thead><tr><th>区域</th><th class="ta-r">中水用量</th><th class="ta-r">占比</th><th class="ta-r">分摊补水</th></tr></thead><tbody>
    ${rc.zones.map(z => `<tr><td>${esc(z.zone)}</td><td class="ta-r">${num(z.usage)}</td><td class="ta-r">${z.share}%</td><td class="ta-r"><b>${num(z.alloc)}</b></td></tr>`).join('')}</tbody></table></div>` : ''}</div></div>` : ''}
  <div class="card"><div class="card-h"><h3>异常提示</h3><span class="hint">共 ${(d.anomalies || []).length} 条</span><span class="sp"></span><a class="btn sm" onclick="go('analysis')">全部</a></div>
    <div class="detail-list">${(d.anomalies || []).length ? d.anomalies.slice(0, 5).map(x => `<div class="it"><div class="t"><span><span class="tag ${x.level === '高' ? 'hi' : x.level === '中' ? 'mid' : 'lo'}">${x.level}</span> ${esc(x.type)} · ${esc(x.meter)}</span></div><div class="d">${esc(x.detail)}</div></div>`).join('') : '<div class="empty">本月未发现异常</div>'}</div></div>
  <div class="card"><div class="card-h"><h3>节能空间测算</h3><span class="hint">理论可节降</span></div>
    <div class="card-b">${(d.savings || []).length ? `<div class="mini"><div><div class="k">可节${isE_() ? '电' : '水'}量</div><div class="v" style="color:#15803d">${num(d.savings.reduce((a, b) => a + b.save_usage, 0))} ${u}</div></div>
    <div><div class="k">可节约费用</div><div class="v" style="color:#b45309">¥${money(d.savings.reduce((a, b) => a + b.save_cost, 0))}</div></div></div>
    ${d.savings.slice(0, 4).map(s => `<div class="bar-row" style="display:flex;justify-content:space-between"><span style="flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(s.title)}</span><b style="color:#15803d">¥${money(s.save_cost)}</b></div>`).join('')}
    <a class="btn sm blk" onclick="go('analysis')">查看完整建议</a>` : '<div class="empty">本月用量平稳，暂无突出节能点</div>'}</div></div>`;
  setTimeout(() => {
    const tr = d.trend || [], map = {}; tr.forEach(x => map[x.month] = x);
    const us = [], ds = [];
    for (let i = 1; i <= 12; i++) { const m = map[S.year + '-' + String(i).padStart(2, '0')]; us.push(m ? m.usage : null); ds.push(m ? m.daily * 30 : null); }
    drawChart($('#chYear'), { labels: Array.from({ length: 12 }, (_, i) => (i + 1) + '月'), series: [{ name: '用量(' + u + ')', data: us, type: 'bar', color: isE_() ? C_ELEC : C_WATER }, { name: '日均×30', data: ds, type: 'line', color: '#f59e0b' }], height: 230 });
    if (groups.length) {
      const ms = groups[0].months.map(m => m.month);
      drawChart($('#chGroup'), { labels: ms.map(m => m.slice(5) + '月'), height: 200,
        series: groups.map((g, i) => ({ name: g.name, type: 'bar', color: ckColor(g.color) || PALETTE[i % PALETTE.length], data: g.months.map(m => m.usage) })) });
    }
  }, 30);
}
function setStructBy(v) { S.structBy = v; S.catPick = []; $('#structBox').innerHTML = structBox(S.dash); }
function toggleCat(id) { const i = S.catPick.indexOf(id); i >= 0 ? S.catPick.splice(i, 1) : S.catPick.push(id); $('#structBox').innerHTML = structBox(S.dash); }
function clearCat() { S.catPick = []; $('#structBox').innerHTML = structBox(S.dash); }
function structBox(d) {
  const byCat = S.structBy === 'cat';
  const list = byCat ? (d.cat || []) : (d.struct || []);
  if (!list.length) return '<div class="empty">暂无数据，请先维护表计分类并录入读数</div>';
  const picked = S.catPick.length ? list.filter(x => S.catPick.includes(x.id)) : list;
  const shown = picked.slice(0, 8), rest = picked.slice(8).reduce((a, b) => a + b.usage, 0);
  const items = rest > 0 ? shown.concat([{ label: '其他' + (picked.length - 8) + '项', usage: rest }]) : shown;
  const total = items.reduce((a, b) => a + b.usage, 0) || 1;
  const tv = (d.cmp || {}).cur || 0, sv = d.struct_sum || (d.struct || []).reduce((a, x) => a + x.usage, 0), diff = sv - tv;
  return `${byCat ? `<div class="chips" style="padding:0 0 8px;border-bottom:1px dashed #eef2f7;margin-bottom:8px">${list.map(x => `<span class="chip ${!S.catPick.length || S.catPick.includes(x.id) ? 'on' : ''}" onclick="toggleCat(${x.id})">${esc(x.label.length > 8 ? x.label.slice(0, 8) : x.label)}</span>`).join('')}${S.catPick.length ? `<span class="chip" onclick="clearCat()">全选</span>` : ''}</div>` : ''}
  <div style="display:flex;gap:12px;align-items:center">${donut(items)}
    <div style="flex:1;min-width:0">${items.map((it, i) => `<div style="display:flex;align-items:center;gap:6px;font-size:11.5px;margin-bottom:5px">
      <i style="width:8px;height:8px;border-radius:2px;flex:0 0 8px;background:${PALETTE[i % PALETTE.length]}"></i>
      <span style="flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(it.label)}</span>
      <b>${(it.usage / total * 100).toFixed(1)}%</b></div>`).join('')}</div></div>
  <div style="margin-top:10px;padding:8px 10px;background:#f8fafc;border-radius:8px;font-size:11px;color:#5b7189;line-height:1.7">
    <b>口径：</b>本图按<b>分表</b>统计，合计 ${num(sv)} ${u_()}；总表为 ${num(tv)} ${u_()}。
    ${Math.abs(diff) > 0.5 ? `差额 <b style="color:${diff > 0 ? '#b91c1c' : '#15803d'}">${diff > 0 ? '+' : ''}${num(diff)}</b>，${diff > 0 ? '分表大于总表，可能重复计量，需核对台账层级' : '属线路损耗与未计量部分，正常'}。` : '两者一致。'} 与总表口径不可相加。</div>`;
}

/* ================= 2. 深度分析 ================= */
async function viewAnalysis(el) {
  const d = await get('analysis', { kind: S.kind, month: S.month, ...zoneQS() });
  const u = u_(), c = d.cmp || {};
  el.innerHTML = `
  <div class="card"><div class="card-h"><h3>核心指标</h3><span class="hint">${fmtMonth(d.month)}</span></div>
    <div class="card-b flush"><div class="tbl-wrap"><table><thead><tr><th>指标</th><th class="ta-r">本月</th><th class="ta-r">上月</th><th class="ta-r">环比</th><th class="ta-r">去年同期</th><th class="ta-r">同比</th></tr></thead><tbody>
    <tr><td><b>${kindName(S.kind)}量(${u})</b></td><td class="ta-r"><b>${num(c.cur)}</b></td><td class="ta-r">${num(c.prev)}</td><td class="ta-r">${pctTag(c.mom)}</td><td class="ta-r">${num(c.last_year)}</td><td class="ta-r">${c.yoy == null ? '<span class="tag flat">无数据</span>' : pctTag(c.yoy)}</td></tr>
    <tr><td>费用(元)</td><td class="ta-r"><b>${money(c.cost)}</b></td><td class="ta-r">—</td><td class="ta-r">—</td><td class="ta-r">—</td><td class="ta-r">—</td></tr>
    <tr><td>日均(${u}/天)</td><td class="ta-r">${num(c.daily_cur, 2)}</td><td class="ta-r">${num(c.daily_prev, 2)}</td><td class="ta-r">${pctTag(c.mom_daily)}</td><td class="ta-r">—</td><td class="ta-r">—</td></tr>
    </tbody></table></div></div></div>
  <div class="notice">同比需去年同月数据，补录上年同期读数后自动显示。</div>
  <div class="card"><div class="card-h"><h3>分系统结构</h3><span class="hint">谁是大头、谁在涨</span></div>
    <div class="card-b flush"><div class="tbl-wrap"><table><thead><tr><th>系统</th><th class="ta-r">用量</th><th class="ta-r">占比</th><th class="ta-r">上月</th><th class="ta-r">环比</th><th class="ta-r">同比</th><th class="ta-r">费用</th></tr></thead><tbody>
    ${(d.struct || []).map(i => `<tr><td>${esc(i.label)}</td><td class="ta-r">${num(i.usage)}</td><td class="ta-r"><b>${i.share}%</b></td><td class="ta-r">${num(i.prev)}</td><td class="ta-r">${pctTag(i.mom)}</td><td class="ta-r">${pctTag(i.yoy)}</td><td class="ta-r">${money(i.cost)}</td></tr>`).join('')}
    <tr class="sum"><td>合计</td><td class="ta-r">${num((d.struct || []).reduce((a, b) => a + b.usage, 0))}</td><td class="ta-r">100%</td><td></td><td></td><td></td><td class="ta-r">${money((d.struct || []).reduce((a, b) => a + (b.cost || 0), 0))}</td></tr>
    </tbody></table></div></div></div>
  <div class="card"><div class="card-h"><h3>分区域结构</h3></div>
    <div class="card-b flush"><div class="tbl-wrap"><table><thead><tr><th>区域</th><th class="ta-r">用量</th><th class="ta-r">占比</th><th class="ta-r">环比</th><th class="ta-r">费用</th></tr></thead><tbody>
    ${(d.zone || []).map(i => `<tr><td>${esc(i.label)}</td><td class="ta-r">${num(i.usage)}</td><td class="ta-r"><b>${i.share}%</b></td><td class="ta-r">${pctTag(i.mom)}</td><td class="ta-r">${money(i.cost)}</td></tr>`).join('')}</tbody></table></div></div></div>
  <div class="card"><div class="card-h"><h3>系统逐月趋势</h3></div><div class="card-b"><div id="chSt" class="chart mid"></div></div></div>
  <div class="card"><div class="card-h"><h3>用量排名 TOP 20</h3></div>
    <div class="card-b flush"><div class="tbl-wrap"><table><thead><tr><th>#</th><th>表计</th><th>系统</th><th>区域</th><th class="ta-r">本月</th><th class="ta-r">环比</th><th class="ta-r">费用</th></tr></thead><tbody>
    ${(d.top || []).map((t, i) => `<tr><td>${i + 1}</td><td><span class="ell">${esc(t.name)}</span>${t.is_total ? ' <span class="tag lo">总</span>' : ''}</td><td>${esc(t.sys)}</td><td>${esc(t.zone || '—')}</td><td class="ta-r"><b>${num(t.usage)}</b></td><td class="ta-r">${pctTag(t.mom)}</td><td class="ta-r">${money(t.cost)}</td></tr>`).join('')}
    </tbody></table></div></div></div>
  <div class="card"><div class="card-h"><h3>异常与风险</h3><span class="hint">共 ${(d.anomalies || []).length} 条</span></div>
    <div class="detail-list">${(d.anomalies || []).length ? d.anomalies.map(x => `<div class="it"><div class="t"><span><span class="tag ${x.level === '高' ? 'hi' : x.level === '中' ? 'mid' : 'lo'}">${x.level}</span> ${esc(x.type)}</span><span style="font-size:11px;color:#8497a8">${esc(x.meter)}</span></div><div class="d">${esc(x.detail)}｜${esc(x.impact)}</div></div>`).join('') : '<div class="empty">本月未发现异常</div>'}</div></div>
  <div class="card"><div class="card-h"><h3>节能建议与测算</h3><span class="hint">基于本月结构自动生成</span></div>
    <div class="detail-list">${(d.savings || []).length ? d.savings.map(s => `<div class="it"><div class="t"><span>${s.is_sum ? '🧮 ' : ''}${esc(s.title)}</span><b style="color:#15803d">${s.is_sum ? '' : '¥'}${money(s.save_cost)}</b></div>
      <div class="d">${esc(s.advice)}</div><div class="d">依据：${esc(s.basis)}｜预计节降 ${num(s.save_usage)} ${u}（${s.ratio}%）</div></div>`).join('') : '<div class="empty">暂无突出节能点</div>'}</div></div>
  <div class="card"><div class="card-h"><h3>抄表覆盖</h3></div><div class="card-b">
    <div class="mini"><div><div class="k">应有分表</div><div class="v">${(d.cover || {}).total}</div></div><div><div class="k">已有数据</div><div class="v">${(d.cover || {}).covered}</div></div><div><div class="k">覆盖率</div><div class="v">${(d.cover || {}).pct}%</div></div></div></div></div>`;
  setTimeout(() => {
    const st = d.struct_trend || [];
    drawChart($('#chSt'), { labels: (d.months || []).slice(-12).map(m => m.slice(2)), height: 210,
      series: st.slice(0, 6).map((s, i) => ({ name: s.label, type: 'line', color: PALETTE[i % PALETTE.length], data: s.data.slice(-12) })) });
  }, 30);
}

/* ================= 3. 趋势 ================= */
async function viewTrend(el) {
  const months = [];
  for (let i = 11; i >= 0; i--) { const d = new Date(S.year, S.month.slice(5) - 1 - i, 1); months.push(d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0')); }
  const [tr, gr] = await Promise.all([
    get('trend', { kind: S.kind, ...zoneQS() }),
    get('groups_trend', { kind: S.kind, months: months.join(',') }),
  ]);
  const list = (tr.months || tr || []).filter(x => x && x.month && months.includes(x.month));
  const u = u_();
  el.innerHTML = `
  <div class="card"><div class="card-h"><h3>近 12 个月用量与日均</h3><span class="hint">柱=用量 线=日均</span></div>
    <div class="card-b"><div id="c1" class="chart tall"></div></div></div>
  <div class="card"><div class="card-h"><h3>费用趋势</h3></div><div class="card-b"><div id="c2" class="chart mid"></div></div></div>
  <div class="card"><div class="card-h"><h3>月度明细</h3></div>
    <div class="card-b flush"><div class="tbl-wrap"><table><thead><tr><th>月份</th><th class="ta-r">用量(${u})</th><th class="ta-r">日均</th><th class="ta-r">费用(元)</th><th class="ta-r">环比</th></tr></thead><tbody>
    ${list.map((x, i) => { const p = list[i - 1]; return `<tr${x.month === S.month ? ' style="background:#eff6ff;font-weight:700"' : ''}><td>${x.month}</td><td class="ta-r">${num(x.usage)}</td><td class="ta-r">${num(x.daily, 2)}</td><td class="ta-r">${money(x.cost)}</td><td class="ta-r">${pctTag(p && p.usage ? (x.usage - p.usage) / p.usage * 100 : null)}</td></tr>`; }).join('')}
    </tbody></table></div></div></div>
  ${(gr.groups || []).length ? `<div class="card"><div class="card-h"><h3>分组对比</h3><span class="hint">在系统设置中维护</span></div>
    <div class="card-b"><div id="c3" class="chart mid"></div></div></div>` : ''}`;
  setTimeout(() => {
    const lb = list.map(x => x.month.slice(2)), c = isE_() ? C_ELEC : C_WATER;
    drawChart($('#c1'), { labels: lb, height: 230, series: [{ name: '用量', data: list.map(x => x.usage), type: 'bar', color: c }, { name: '日均', data: list.map(x => x.daily), type: 'line', color: '#f59e0b' }] });
    drawChart($('#c2'), { labels: lb, height: 190, series: [{ name: '费用(元)', data: list.map(x => x.cost), type: 'bar', color: '#16a34a' }] });
    drawChart($('#c3'), { labels: months.map(m => m.slice(2)), height: 200, series: (gr.groups || []).map((g, i) => ({ name: g.name, type: 'bar', color: ckColor(g.color) || PALETTE[i % PALETTE.length], data: g.months.map(m => m.usage) })) });
  }, 30);
}

/* ================= 4. 损耗分析 ================= */
async function viewLoss(el) {
  const months = [];
  for (let i = 11; i >= 0; i--) { const d = new Date(S.year, S.month.slice(5) - 1 - i, 1); months.push(d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0')); }
  const l = await get('loss', { kind: S.kind, months: months.join(','), ...zoneQS() });
  const u = u_(), rows = l.rows || [], E = isE_();
  el.innerHTML = `
  <div class="notice">${E ? '①国网关口表用于核对自抄误差；②一级总表为用量口径；③分表合计衡量未计量部分。三者不可相加。' : '①自来水总表为用量口径；②分表合计用于校核损耗。'}</div>
  <div class="card"><div class="card-h"><h3>三级总表比对</h3><span class="hint">近 12 个月</span></div>
    <div class="card-b"><div id="cLoss" class="chart tall"></div></div></div>
  <div class="card"><div class="card-h"><h3>损耗明细</h3></div>
    <div class="card-b flush"><div class="tbl-wrap"><table><thead><tr><th>月份</th>
    ${E ? '<th class="ta-r">①国网</th>' : ''}<th class="ta-r">${E ? '②一级总表' : '①总表'}</th><th class="ta-r">${E ? '③分表' : '②分表'}</th>
    <th class="ta-r">${E ? '变损线损' : '损耗'}</th><th class="ta-r">损耗率</th><th class="ta-r">未计量</th></tr></thead><tbody>
    ${rows.map(r => `<tr><td>${r.month}</td>${E ? `<td class="ta-r">${num(r.g)}</td>` : ''}<td class="ta-r"><b>${num(r.l1)}</b></td><td class="ta-r">${num(r.sub)}</td>
    <td class="ta-r">${num(E ? r.d_g_l1 : r.d_l1_sub)}</td><td class="ta-r">${E ? (r.d_g_l1_pct == null ? '—' : r.d_g_l1_pct + '%') : (r.d_l1_sub_pct == null ? '—' : r.d_l1_sub_pct + '%')}</td>
    <td class="ta-r">${num(r.d_l1_sub)}${E && r.d_l1_sub_pct != null ? ` (${r.d_l1_sub_pct}%)` : ''}</td></tr>`).join('') || '<tr><td colspan="6"><div class="empty">暂无数据</div></td></tr>'}
    </tbody></table></div></div></div>
  <div class="card"><div class="card-h"><h3>国网核对</h3><span class="hint">自抄 vs 供电公司</span></div>
    <div class="card-b">${(l.check || []).some(c => c.grid_usage != null) ? `<div class="tbl-wrap"><table><thead><tr><th>账期</th><th class="ta-r">国网结算</th><th class="ta-r">自抄</th><th class="ta-r">差值</th><th class="ta-r">误差率</th><th>抄表日</th></tr></thead><tbody>
    ${(l.check || []).filter(c => c.grid_usage != null).map(c => `<tr><td>${c.month}</td><td class="ta-r">${num(c.grid_usage)}</td><td class="ta-r">${num(c.mine)}</td><td class="ta-r">${num(c.diff)}</td><td class="ta-r">${c.gap == null ? '—' : c.gap.toFixed(2) + '%'}</td><td>${esc(c.rdate || '')}</td></tr>`).join('')}</tbody></table></div>`
    : '<div class="empty">未录入国网抄表，点击右下按钮逐月录入</div>'}
    <button class="btn sm blk" onclick="editGrid('${S.month}')">＋ 录入/修改国网 ${S.month}</button></div></div>
  ${(l.detail || []).length ? `<div class="card"><div class="card-h"><h3>各级总表逐月</h3></div>
    <div class="card-b flush"><div class="tbl-wrap"><table><thead><tr><th>表计</th>${months.map(m => `<th class="ta-r">${m.slice(5)}</th>`).join('')}</tr></thead><tbody>
    ${l.detail.map(dt => `<tr><td><span class="ell">${esc((dt.grid ? '【国网】' : dt.level1 ? '【一级】' : '') + dt.name)}</span></td>${dt.data.map(v => `<td class="ta-r">${num(v)}</td>`).join('')}</tr>`).join('')}</tbody></table></div></div></div>` : ''}
  ${l.kind === 'water' && l.reclaimed ? `<div class="card"><div class="card-h"><h3>中水逐月</h3></div>
    <div class="card-b flush"><div class="tbl-wrap"><table><thead><tr><th>月份</th><th class="ta-r">补水</th><th class="ta-r">中水用量</th><th class="ta-r">损耗</th></tr></thead><tbody>
    ${(l.reclaimed.by_month || []).map(r => `<tr><td>${r.month}</td><td class="ta-r">${num(r.supply)}</td><td class="ta-r">${num(r.reclaimed)}</td><td class="ta-r">${num(r.loss)}</td></tr>`).join('')}</tbody></table></div></div></div>` : ''}`;
  setTimeout(() => {
    const ser = [{ name: E ? '②一级总表' : '①总表', data: rows.map(r => r.l1), type: 'bar', color: isE_() ? C_ELEC : C_WATER }, { name: '③/②分表', data: rows.map(r => r.sub), type: 'bar', color: '#94a3b8' }];
    if (E) ser.unshift({ name: '①国网', data: rows.map(r => r.g), type: 'line', color: '#f59e0b' });
    drawChart($('#cLoss'), { labels: rows.map(r => r.month.slice(2)), height: 230, series: ser });
  }, 30);
  }
function editGrid(mk) {
  modal('录入国网抄表 ' + mk, `
    <div class="field"><label>户号/表号</label><input class="inp" id="gno"></div>
    <div class="field"><label>国网抄表日</label><input class="inp" type="date" id="grd"></div>
    <div class="field"><label>国网表底数</label><input class="inp" type="number" id="gval"></div>
    <div class="field"><label>结算${isE_() ? '电量(度)' : '水量(吨)'}</label><input class="inp" type="number" id="guse"></div>
    <div class="field"><label>费用(元)</label><input class="inp" type="number" id="gfee"></div>
    <div class="field"><label>备注</label><input class="inp" id="gnt"></div>`,
    `<button class="btn" onclick="closeModal()">取消</button><button class="btn pri" onclick="saveGrid('${mk}')">保存</button>`);
  get('grid').then(gs => { const g = (gs || []).find(x => x.mk === mk && x.kind === S.kind); if (g) { gno.value = g.meter_no || ''; grd.value = g.rdate || ''; gval.value = g.val ?? ''; guse.value = g.usage ?? ''; gfee.value = g.fee ?? ''; gnt.value = g.note || ''; } });
}
async function saveGrid(oldMk) {
  await post('grid_save', { kind: S.kind, mk: oldMk, meter_no: gno.value, rdate: grd.value, val: gval.value === '' ? null : +gval.value, usage: guse.value === '' ? null : +guse.value, fee: gfee.value === '' ? null : +gfee.value, note: gnt.value });
  closeModal(); toast('已保存', 'ok'); render();
}

/* ================= 5. 表计台账 ================= */
async function viewMeters(el) {
  const ms = await get('meters', S.kind);
  const by = {};
  ms.forEach(m => { (by[m.sys] = by[m.sys] || []).push(m); });
  const keys = Object.keys(by);
  el.innerHTML = `
  <div class="card"><div class="card-h"><h3>${kindName(S.kind)}表计</h3><span class="hint">共 ${ms.length} 块 · 启用 ${ms.filter(m => m.enabled).length}</span></div>
    <div class="card-b">${ms.length ? `<div class="mini"><div><div class="k">总表</div><div class="v">${ms.filter(m => m.is_total).length}</div></div><div><div class="k">二级总表</div><div class="v">${ms.filter(m => m.is_sub_total).length}</div></div><div><div class="k">关口表</div><div class="v">${ms.filter(m => m.is_grid).length}</div></div><div><div class="k">备用</div><div class="v">${ms.filter(m => m.use_status === 'backup').length}</div></div></div>` : '<div class="empty">暂无表计</div>'}
    <div class="btnrow" style="margin-top:10px"><button class="btn pri sm" onclick="editMeter(0,'${S.kind}')">＋ 新增表计</button>
    <button class="btn sm" onclick="dl('export_ledger','表计台账.xlsx')">导出台账</button>
    <button class="btn sm" onclick="pickImport('ledger')">导入台账</button></div></div></div>
  ${keys.map(k => `<div class="card"><div class="card-h"><h3>${esc(k)}</h3><span class="hint">${by[k].length} 块</span></div>
    <div class="card-b flush"><div class="tbl-wrap"><table><thead><tr><th>名称</th><th>区域</th><th>倍率</th><th>层级</th><th>状态</th><th></th></tr></thead><tbody>
    ${by[k].map(m => `<tr style="${m.enabled ? '' : 'opacity:.45'}"><td><span class="ell" style="max-width:130px">${esc(m.name)}</span><br><small style="color:#8497a8">${esc(m.code || '')}${m.loc ? ' · ' + esc(m.loc) : ''}</small></td>
      <td>${esc(m.zone || '—')}</td><td class="ta-r">${m.mult}</td><td>${lvTag(m)}</td>
      <td>${m.use_status === 'backup' ? '<span class="tag">备用</span>' : (m.enabled ? '<span class="tag lo">在用</span>' : '<span class="tag">停用</span>')}${m.source === 'reclaimed' ? '<span class="tag mid">中水</span>' : ''}</td>
      <td><button class="btn sm" onclick="editMeter(${m.id},'${S.kind}')">编辑</button></td></tr>`).join('')}
    </tbody></table></div></div></div>`).join('')}`;
}
function lvTag(m) { if (m.is_grid) return '<span class="tag mid">关口</span>'; if (m.is_total) return '<span class="tag hi">一级</span>'; if (m.is_sub_total) return '<span class="tag up" style="background:#ffedd5;color:#c2410c">二级</span>'; return '<span class="tag flat">分表</span>'; }
function editMeter(id, kind) {
  const m = (S.boot.meters || []).find(x => x.id === id) || { kind, mult: 1, factor: 1, is_total: 0, is_sub_total: 0, is_grid: 0, enabled: 1, use_status: 'main', source: 'tap', sort: 0 };
  const b = S.boot;
  const sysl = kind === 'elec' ? (b.elec_sys || []) : (b.water_sys || []);
  modal(id ? '编辑表计' : '新增表计', `
    <div class="field"><label>表名称/位置 *</label><input class="inp" id="mn" value="${esc(m.name || '')}"></div>
    <div class="field"><label>表编号</label><input class="inp" id="mc" value="${esc(m.code || '')}"></div>
    <div class="field"><label>系统分类</label><select class="inp" id="ms">${sysl.map(s => `<option ${s === m.sys ? 'selected' : ''}>${esc(s)}</option>`).join('')}</select></div>
    <div class="field"><label>区域</label><select class="inp" id="mz">${[''].concat(b.zones || []).map(z => `<option value="${esc(z)}" ${z === (m.zone || '') ? 'selected' : ''}>${esc(z || '（未分区域）')}</option>`).join('')}</select></div>
    <div class="field"><label>安装位置</label><input class="inp" id="ml" value="${esc(m.loc || '')}"></div>
    <div class="field"><label>配电柜/面板</label><input class="inp" id="mp" value="${esc(m.panel || '')}"></div>
    <div class="field"><label>用途</label><input class="inp" id="mq" value="${esc(m.purpose || '')}"></div>
    <div class="field"><label>倍率</label><input class="inp" type="number" step="0.01" id="mm" value="${m.mult ?? 1}"></div>
    <div class="field"><label>折算系数</label><input class="inp" type="number" step="0.01" id="mf" value="${m.factor ?? 1}"></div>
    <div class="field"><label>计量层级</label><select class="inp" id="mlv">
      <option value="0" ${!m.is_total && !m.is_sub_total && !m.is_grid ? 'selected' : ''}>分表</option>
      <option value="t" ${m.is_total ? 'selected' : ''}>一级总表</option>
      <option value="s" ${m.is_sub_total ? 'selected' : ''}>二级总表</option>
      <option value="g" ${m.is_grid ? 'selected' : ''}>国网关口表</option></select></div>
    ${kind === 'water' ? `<div class="field"><label>水源</label><select class="inp" id="msrc"><option value="tap" ${m.source !== 'reclaimed' ? 'selected' : ''}>自来水</option><option value="reclaimed" ${m.source === 'reclaimed' ? 'selected' : ''}>中水</option></select></div>` : ''}
    <div class="field"><label>主备状态</label><select class="inp" id="mus"><option value="main" ${m.use_status !== 'backup' ? 'selected' : ''}>主用</option><option value="backup" ${m.use_status === 'backup' ? 'selected' : ''}>备用（不计入合计）</option></select></div>
    <div class="field"><label>启用</label><select class="inp" id="men"><option value="1" ${m.enabled ? 'selected' : ''}>启用</option><option value="0" ${!m.enabled ? 'selected' : ''}>停用</option></select></div>
    ${id ? '<button class="btn blk" style="color:#b91c1c;border-color:#fecaca" onclick="delMeter(' + id + ',\'' + esc(m.name).replace(/'/g, '') + '\')">删除此表计（含全部读数）</button>' : ''}`,
    `<button class="btn" onclick="closeModal()">取消</button><button class="btn pri" onclick="saveMeter(${id || 0},'${kind}')">保存</button>`);
}
async function saveMeter(id, kind) {
  const lv = mlv.value;
  const d = { id: id || undefined, kind, name: mn.value.trim(), code: mc.value.trim(), sys: ms.value, zone: mz.value, loc: ml.value, panel: mp.value, purpose: mq.value,
    mult: +mm.value || 1, factor: +mf.value || 1, use_status: mus.value, enabled: +men.value, is_total: lv === 't' ? 1 : 0, is_sub_total: lv === 's' ? 1 : 0, is_grid: lv === 'g' ? 1 : 0 };
  if (msrc) d.source = msrc.value;
  if (!d.name) return toast('请填写表名称', 'err');
  await post('meter', d); closeModal(); toast('已保存', 'ok');
  S.boot = await get('bootstrap'); render();
}
function delMeter(id, name) {
  if (!confirm('确定删除「' + name + '」？该表全部读数将一并删除，不可恢复！')) return;
  post('meter_delete', { id, name }).then(() => { closeModal(); toast('已删除', 'ok'); return loadAll(); });
}

/* ================= 6. 抄表录入 ================= */
async function viewReading(el) {
  const mk = S.month;
  const r = await get('reading_matrix', { month: mk, kind: S.kind });
  S.rd = { mk, rows: (r.rows || []).map(x => ({ ...x, input: x.val != null && x.rdate === mk + '-' + String(new Date(mk.slice(0, 4), mk.slice(5), 0).getDate()).padStart(2, '0') ? String(x.val) : '', base: x.val != null ? String(x.val) : '' })) };
  S.rdDirty = false;
  el.innerHTML = `
  <div class="notice">录入 <b>${mk}</b> 的<b>最新表底读数</b>，系统按前后读数自动计算用量并按天归月。只需填本月有新读数的表。</div>
  <div class="card"><div class="card-h"><h3>${isE_() ? '⚡ 电表' : '💧 水表'} · ${S.rd.rows.length} 块</h3>
    <span class="sp"></span><button class="btn sm" onclick="rdQuick('all')">全部复制上次</button></div>
    <div class="card-b" id="rdList">${rdRows()}</div>
    <div class="rd-save" style="padding:0 12px 12px"><button class="btn pri blk" onclick="saveReadings()">保存本次抄表</button></div></div>`;
}
function rdRows() {
  const q = (rdQ || '').trim();
  return `<input class="inp sm" placeholder="🔍 搜索表名/编号/区域" oninput="rdQ=this.value;$('#rdList').innerHTML=rdRows()" style="margin-bottom:6px" value="${esc(q)}">` +
    S.rd.rows.filter(x => !q || (x.name + x.code + x.zone + x.sys).includes(q)).map(x => {
      const i = S.rd.rows.indexOf(x);
      return `<div class="rd-row ${x.input !== '' ? 'reading' : ''}"><div class="info"><div class="n">${esc(x.name)}</div>
      <div class="m">${esc(x.sys)}${x.zone ? ' · ' + esc(x.zone) : ''}${x.total ? ' · <b>总表</b>' : ''}</div></div>
      <div class="prev">上次<br><b>${x.prev_val == null ? '—' : num(x.prev_val, 0)}</b><br><small>${esc((x.prev_date || '').slice(5) || '')}</small></div>
      <input class="inp" type="number" inputmode="decimal" placeholder="—" value="${esc(x.input)}" oninput="rdInput(${i},this.value)"></div>`;
    }).join('') || '<div class="empty">无匹配表计</div>';
}
let rdQ = '';
function rdInput(i, v) { S.rd.rows[i].input = v; S.rdDirty = true; }
function rdQuick() { S.rd.rows.forEach(x => { if (!x.input && x.prev_val != null) x.input = x.base; }); S.rdDirty = true; $('#rdList').innerHTML = rdRows(); }
async function saveReadings() {
  const rows = [];
  S.rd.rows.forEach(x => { if (x.input !== '' && x.input != null) rows.push({ meter_id: x.id, rdate: S.rd.mk + '-' + String(new Date(S.rd.mk.slice(0, 4), S.rd.mk.slice(5), 0).getDate()).padStart(2, '0'), val: +x.input }); });
  if (!rows.length) return toast('还没有填写任何读数', 'err');
  await post('readings_batch', { rows });
  S.rdDirty = false; toast('已保存 ' + rows.length + ' 条读数', 'ok');
  S.dash = null; render();
}

/* ================= 7. 数据导入 ================= */
async function viewImport(el) {
  el.innerHTML = `
  <div class="card"><div class="card-h"><h3>Excel 抄表导入</h3></div><div class="card-b">
    <div style="font-size:12px;color:#8497a8;line-height:1.7;margin-bottom:10px">表头第一行为日期（如 2026.3.31），每行一块表；工作表名含「水」按水处理，否则按电。系统自动识别表编号、名称、配电柜、用途并生成/更新台账。</div>
    <button class="btn pri blk" onclick="pickImport('readings')">选择 Excel 文件导入</button>
    <button class="btn blk" onclick="dl('export_template','抄表导入模板.xlsx')">下载导入模板</button>
    <button class="btn blk" onclick="dl('export_fill_template','抄表填空模板.xlsx')">下载填空模板（只填数字）</button></div></div>
  <div class="card"><div class="card-h"><h3>表计台账导入</h3></div><div class="card-b">
    <div style="font-size:12px;color:#8497a8;line-height:1.7;margin-bottom:10px">按导出的台账格式导入，表头需含「表名称」「类型」等列；以表编号为唯一键更新。</div>
    <button class="btn blk" onclick="pickImport('ledger')">选择台账 Excel</button></div></div>
  <div class="card"><div class="card-h"><h3>读数明细导入</h3></div><div class="card-b">
    <div style="font-size:12px;color:#8497a8;line-height:1.7;margin-bottom:10px">列：表编号 / 表名称 / 类型 / 日期 / 读数。按编号或名称匹配已有表。</div>
    <button class="btn blk" onclick="pickImport('readings_detail')">选择读数明细 Excel</button></div></div>
  <input type="file" id="impFile" accept=".xlsx" style="display:none" onchange="doImport(this)">`;
}
function pickImport(t) { S.impType = t; impFile.value = ''; impFile.click(); }
async function dlLedger() { dl('export_ledger', '表计台账.xlsx'); }
function doImport(inp) {
  const f = inp.files[0]; if (!f) return;
  const rd = new FileReader();
  rd.onload = async () => {
    const b64 = btoa(new Uint8Array(rd.result).reduce((s, c) => s + String.fromCharCode(c), ''));
    toast('导入中…');
    try {
      if (S.impType === 'readings') { const r = await post('import_xlsx', { content: b64, filename: f.name }); showImp(r); }
      else if (S.impType === 'ledger') { const r = await post('import_ledger', { content: b64 }); toast(`台账导入：新增 ${r.added} 更新 ${r.updated}`, 'ok'); loadAll(); }
      else { const r = await post('import_readings', { content: b64 }); toast(`读数导入：${r.reads} 条，未匹配 ${r.missed}`, 'ok'); }
    } catch (e) { toast('导入失败：' + e.message, 'err'); }
  };
  rd.readAsArrayBuffer(f);
}
function showImp(r) {
  const s = (r.sheets || []).map(x => `<tr><td>${esc(x.sheet)}</td><td>${x.kind === 'water' ? '水' : '电'}</td><td>${x.ok ? `新增${x.added} 更新${x.updated} 读数${x.reads}` : '<span style="color:#b91c1c">' + esc(x.error) + '</span>'}</td></tr>`).join('');
  modal('导入结果', `<div class="tbl-wrap"><table><thead><tr><th>工作表</th><th>类型</th><th>结果</th></tr></thead><tbody>${s}</tbody></table></div>`);
  S.dash = null;
}

/* ================= 8. 报表导出 ================= */
async function viewExport(el) {
  const mt = S.month.slice(5) | 0;
  el.innerHTML = `
  <div class="card"><div class="card-h"><h3>报表导出</h3><span class="hint">${fmtMonth(S.month)}</span></div><div class="card-b">
    <div style="font-size:12px;color:#8497a8;margin-bottom:10px">在手机上点击后生成文件，浏览器会下载到「下载/Download」目录；也可在电脑打开同一地址导出。</div>
    <div class="btnrow">
      <button class="btn pri" onclick="dl('export_report?kind=${S.kind}&month=${S.month}&year=${S.year}','能耗分析报表.xlsx')">📊 ${kindName(S.kind)}分析报表(Excel)</button>
      <button class="btn" onclick="dl('export_data','能耗数据明细.xlsx')">🗂 全量数据明细</button>
      <button class="btn" onclick="dl('export_ledger','表计台账.xlsx')">▤ 表计台账</button>
      <button class="btn" onclick="dl('export_readings','抄表读数明细.xlsx')">✎ 抄表读数明细</button>
      <button class="btn" onclick="dl('export_fill_template?month=${S.month}','抄表填空模板.xlsx')">☐ ${mt}月填空模板</button>
      <button class="btn" onclick="dl('export_template','抄表导入模板.xlsx')">⇪ 导入模板</button>
    </div></div></div>
  <div class="card"><div class="card-h"><h3>数据备份 / 恢复</h3></div><div class="card-b">
    <div style="font-size:12px;color:#8497a8;margin-bottom:10px">备份为 JSON 文件，换手机或重装后可恢复。</div>
    <div class="btnrow"><button class="btn" onclick="dl('backup','能耗数据备份.json')">⬇ 导出备份</button>
    <button class="btn" onclick="pickRestore()">⬆ 恢复备份</button></div>
    <input type="file" id="resFile" accept=".json" style="display:none" onchange="doRestore(this)"></div></div>`;
}
function dl(path, name) {
  toast('正在生成…');
  get(path).then(d => {
    const a = document.createElement('a');
    a.href = d.url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
  }).catch(e => toast(e.message, 'err'));
}
function pickRestore() { resFile.value = ''; resFile.click(); }
function doRestore(inp) {
  const f = inp.files[0]; if (!f) return;
  if (!confirm('恢复将覆盖当前全部数据，确定继续？')) return;
  const rd = new FileReader();
  rd.onload = () => { post('restore', { content: rd.result }).then(() => { toast('恢复成功', 'ok'); loadAll(); }).catch(e => toast(e.message, 'err')); };
  rd.readAsText(f);
}

/* ================= 9. 抄表质量 ================= */
async function viewQuality(el) {
  const q = await get('quality', { kind: S.kind });
  const bad = (q || []).filter(x => x.status !== '正常' && x.status !== '无数据');
  el.innerHTML = `
  <div class="card"><div class="card-h"><h3>抄表质量画像</h3><span class="hint">${kindName(S.kind)} · 共 ${(q || []).length} 块表</span></div>
    <div class="card-b">${bad.length ? `<div class="notice bad">${bad.length} 块表存在异常：${esc(bad.slice(0, 5).map(x => x.name + '(' + x.status + ')').join('、'))}${bad.length > 5 ? ' 等' : ''}</div>` : '<div class="notice">全部表计抄表间隔正常</div>'}
    <div style="font-size:11.5px;color:#8497a8">状态判定：最大抄表间隔 &gt; 45 天为「间隔过长」；出现读数倒走单独标记。</div></div></div>
  <div class="card"><div class="card-b flush"><div class="tbl-wrap"><table><thead><tr><th>表计</th><th>系统</th><th class="ta-r">抄表次数</th><th class="ta-r">平均间隔</th><th class="ta-r">最大间隔</th><th>状态</th></tr></thead><tbody>
  ${(q || []).map(x => `<tr><td><span class="ell" style="max-width:140px">${esc(x.name)}</span></td><td>${esc(x.sys || '')}</td><td class="ta-r">${x.cnt}</td><td class="ta-r">${x.avg_gap}</td><td class="ta-r">${x.max_gap}</td>
    <td>${x.status === '正常' ? '<span class="tag lo">正常</span>' : x.status === '无数据' ? '<span class="tag">无数据</span>' : `<span class="tag ${x.status === '读数倒走' ? 'hi' : 'mid'}">${esc(x.status)}</span>`}</td></tr>`).join('')}
  </tbody></table></div></div></div>`;
}

/* ================= 10. 系统设置 ================= */
async function viewSettings(el) {
  const b = S.boot;
  const pmE = b.price_months ? b.price_months.elec || {} : {}, pmW = b.price_months ? b.price_months.water || {} : {};
  const curPm = (k) => { const o = k === 'elec' ? pmE : pmW; return Object.keys(o).sort().reverse().slice(0, 6); };
  el.innerHTML = `
  <div class="card"><div class="card-h"><h3>项目信息</h3></div><div class="card-b">
    <div class="field"><label>项目名称</label><input class="inp" id="stProject" value="${esc(b.meta.project || '')}"></div>
    <div class="field"><label>系统标题</label><input class="inp" id="stTitle" value="${esc(b.meta.title || '')}"></div>
    <button class="btn pri blk" onclick="saveSettings()">保存项目信息</button></div></div>
  <div class="card"><div class="card-h"><h3>默认单价</h3><span class="hint">元/度、元/吨</span></div><div class="card-b">
    <div style="display:flex;gap:8px"><div class="field" style="flex:1"><label>⚡ 电价</label><input class="inp" type="number" step="0.0001" id="prE" value="${num((b.prices || {}).elec, 4)}"></div>
    <div class="field" style="flex:1"><label>💧 水价</label><input class="inp" type="number" step="0.0001" id="prW" value="${num((b.prices || {}).water, 4)}"></div></div>
    <button class="btn pri blk" onclick="savePrices()">保存单价</button></div></div>
  <div class="card"><div class="card-h"><h3>按月单价</h3><span class="hint">电价每月可能不同</span></div><div class="card-b">
    <div style="display:flex;gap:8px;align-items:flex-end"><div class="field" style="flex:1"><label>类型</label><select class="inp" id="pmKind"><option value="elec">⚡ 电</option><option value="water">💧 水</option></select></div>
    <div class="field" style="flex:1"><label>月份</label><input class="inp" type="month" id="pmMk" value="${S.month}"></div>
    <div class="field" style="flex:1"><label>单价</label><input class="inp" type="number" step="0.0001" id="pmV" placeholder="留空=删除"></div>
    <button class="btn pri" style="margin-bottom:10px" onclick="savePriceMonth()">设</button></div>
    <div class="tbl-wrap"><table><thead><tr><th>类型</th><th>月份</th><th class="ta-r">单价</th><th></th></tr></thead><tbody>
    ${['elec', 'water'].flatMap(k => curPm(k).map(mk => `<tr><td>${k === 'elec' ? '⚡' : '💧'}</td><td>${mk}</td><td class="ta-r">${num((k === 'elec' ? pmE : pmW)[mk], 4)}</td>
      <td><button class="btn sm" onclick="delPriceMonth('${k}','${mk}')">删</button></td></tr>`)).join('') || '<tr><td colspan="4"><div class="empty">未设置按月单价，使用默认单价</div></td></tr>'}
    </tbody></table></div></div></div>
  <div class="card"><div class="card-h"><h3>分类字典</h3><span class="hint">有表计在用的分类不能删除</span></div><div class="card-b">
    ${[['sys_elec', '⚡ 系统分类'], ['sys_water', '💧 系统分类'], ['zone', '📍 区域']].map(([k, t]) => `
      <div style="margin-bottom:12px"><div style="font-size:12.5px;font-weight:600;margin-bottom:6px">${t}</div>
      <div class="chips" style="padding:0">${(b.dicts || []).filter(x => x.kind === k).map(x => `<span class="chip on" onclick="delDict(${x.id},'${esc(x.name).replace(/'/g, '')}','${k}')">${esc(x.name)} ×</span>`).join('')}
      <span class="chip" onclick="addDict('${k}')">＋ 新增</span></div></div>`).join('')}
  </div></div>
  <div class="card"><div class="card-h"><h3>显示分组</h3><span class="hint">如 A区+D区+北通道</span></div><div class="card-b">
    ${(b.groups || []).map(g => `<div class="bar-row"><div class="nm" style="width:100px;flex:0 0 100px"><i style="display:inline-block;width:8px;height:8px;border-radius:2px;background:${ckColor(g.color)};margin-right:4px"></i>${esc(g.name)}</div>
      <div style="flex:1;font-size:11px;color:#8497a8;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc((g.zones || []).join('、') || '全部区域')}｜${esc(g.kind)}</div>
      <button class="btn sm" onclick="editGroup(${g.id})">改</button></div>`).join('') || '<div class="empty">暂无分组</div>'}
    <button class="btn sm blk" onclick="editGroup(0)">＋ 新增分组</button></div></div>
  <div class="card"><div class="card-h"><h3>账户安全</h3></div><div class="card-b">
    <div class="field"><label>原密码</label><input class="inp" type="password" id="pwOld"></div>
    <div class="field"><label>新密码（至少 4 位）</label><input class="inp" type="password" id="pwNew"></div>
    <button class="btn pri blk" onclick="changePwd()">修改密码</button></div></div>
  <div style="text-align:center;font-size:11px;color:#8497a8;padding:6px 0 10px">© ${new Date().getFullYear()} 开心团团</div>`;
}
async function saveSettings() { await post('settings', { project: stProject.value, title: stTitle.value }); toast('已保存', 'ok'); loadAll(); }
async function savePrices() { await post('price', { elec: +prE.value, water: +prW.value }); toast('已保存', 'ok'); loadAll(); }
async function savePriceMonth() { if (!pmMk.value) return toast('请选择月份', 'err'); await post('price_month', { kind: pmKind.value, mk: pmMk.value, value: pmV.value === '' ? null : +pmV.value }); toast('已设置', 'ok'); loadAll(); }
async function delPriceMonth(k, mk) { await post('price_month', { kind: k, mk, value: null }); loadAll(); }
function addDict(kind) {
  modal('新增分类', `<div class="field"><label>名称</label><input class="inp" id="dn" placeholder="如 G区"></div>`,
    `<button class="btn" onclick="closeModal()">取消</button><button class="btn pri" onclick="doAddDict('${kind}')">添加</button>`);
}
async function doAddDict(kind) { const n = (dn.value || '').trim(); if (!n) return; try { await post('dict_add', { kind, name: n }); closeModal(); toast('已添加', 'ok'); loadAll(); } catch (e) { toast(e.message, 'err'); } }
async function delDict(id, name, kind) {
  if (!confirm(`删除分类「${name}」？`)) return;
  try { const r = await post('dict_del', { id }); if (!r.ok) throw new Error(r.msg || '无法删除'); toast('已删除', 'ok'); loadAll(); } catch (e) { toast(e.message, 'err'); }
}
function editGroup(id) {
  const g = (S.boot.groups || []).find(x => x.id === id) || { name: '', zones: [], sys: [], kind: 'both', color: '#1a56db' };
  const zones = S.boot.zones || [];
  modal(id ? '编辑分组' : '新增分组', `
    <div class="field"><label>分组名称</label><input class="inp" id="gn" value="${esc(g.name)}"></div>
    <div class="field"><label>包含区域（不选=全部）</label><div class="chips" style="padding:0" id="gz">${zones.map(z => `<span class="chip ${g.zones.includes(z) ? 'on' : ''}" data-z="${esc(z)}" onclick="this.classList.toggle('on')">${esc(z)}</span>`).join('')}</div></div>
    <div class="field"><label>适用类型</label><select class="inp" id="gk"><option value="both" ${g.kind === 'both' ? 'selected' : ''}>电+水</option><option value="elec" ${g.kind === 'elec' ? 'selected' : ''}>仅电</option><option value="water" ${g.kind === 'water' ? 'selected' : ''}>仅水</option></select></div>
    <div class="field"><label>颜色</label><input type="color" class="inp" id="gc" value="${ckColor(g.color)}" style="padding:4px;height:40px"></div>
    ${id ? '<button class="btn blk" style="color:#b91c1c;border-color:#fecaca" onclick="delGroup(' + id + ')">删除此分组</button>' : ''}`,
    `<button class="btn" onclick="closeModal()">取消</button><button class="btn pri" onclick="saveGroup(${id || 0})">保存</button>`);
}
async function saveGroup(id) {
  const zones = Array.from(document.querySelectorAll('#gz .chip.on')).map(c => c.dataset.z);
  await post('group', { id: id || undefined, name: gn.value.trim() || '未命名', zones, sys: [], kind: gk.value, color: gc.value, sort: 99 });
  closeModal(); toast('已保存', 'ok'); loadAll();
}
function delGroup(id) { if (!confirm('删除该分组？')) return; post('group_delete', { id }).then(() => { closeModal(); toast('已删除', 'ok'); loadAll(); }); }
async function changePwd() {
  if ((pwNew.value || '').length < 4) return toast('新密码至少 4 位', 'err');
  try { await post('auth_change', { old: pwOld.value, new: pwNew.value }); toast('密码已修改', 'ok'); pwOld.value = pwNew.value = ''; } catch (e) { toast(e.message, 'err'); }
}

/* ================= 11. 使用帮助 ================= */
async function viewHelp(el) {
  el.innerHTML = `
  <div class="card"><div class="card-h"><h3>快速上手</h3></div><div class="detail-list">
    <div class="it"><div class="t">1. 维护表计台账</div><div class="d">「表计」页新增或导入 Excel，设置系统分类、区域、计量层级（关口/一级/二级/分表）。</div></div>
    <div class="it"><div class="t">2. 录入抄表</div><div class="d">「抄表」页输入每月表底读数；也可在「更多 → 数据导入」批量导入抄表 Excel。</div></div>
    <div class="it"><div class="t">3. 查看分析</div><div class="d">「总览 / 分析 / 趋势 / 损耗」查看同比环比、结构、异常与节能建议。</div></div>
    <div class="it"><div class="t">4. 设置单价</div><div class="d">「更多 → 系统设置」设置默认单价与按月单价，费用自动计算。</div></div>
    <div class="it"><div class="t">5. 导出报表</div><div class="d">「更多 → 报表导出」生成 Excel 分析报表，文件保存到手机「下载」目录。</div></div>
  </div></div>
  <div class="card"><div class="card-h"><h3>计算口径</h3></div><div class="detail-list">
    <div class="it"><div class="t">单表用量</div><div class="d">本次读数 − 上次读数 × 倍率 × 折算系数</div></div>
    <div class="it"><div class="t">用量归月</div><div class="d">跨月抄表按天均摊到各自然月</div></div>
    <div class="it"><div class="t">总用量</div><div class="d">电 = 低压进线柜一级总表合计；水 = 自来水总表合计；无总表月份自动回落到二级总表 / 分表合计</div></div>
    <div class="it"><div class="t">环比 / 同比</div><div class="d">(本月 − 上月) ÷ 上月 × 100%；同比对比去年同月</div></div>
    <div class="it"><div class="t">日均</div><div class="d">用量 ÷ 当月天数，可消除月份天数差异</div></div>
    <div class="it"><div class="t">结构占比</div><div class="d">按分表（末端计量）统计，与总表口径不同，不可相加</div></div>
    <div class="it"><div class="t">变损线损</div><div class="d">国网结算 − 自抄一级总表；未计量 = 一级总表 − 分表合计</div></div>
    <div class="it"><div class="t">备用表</div><div class="d">不计入任何合计；中水表单独统计，不与自来水混合</div></div>
  </div></div>
  <div class="card"><div class="card-h"><h3>安装到手机桌面</h3></div><div class="card-b" style="font-size:12.5px;color:#5b7189;line-height:1.8">
    <b>小米/Chrome 浏览器：</b>打开本系统 → 右上角 ⋮ → 「添加到主屏幕 / 安装应用」<br>
    <b>苹果 Safari：</b>分享按钮 → 「添加到主屏幕」<br>
    添加后从桌面图标打开即为全屏模式，并支持离线查看最近一次的数据。</div></div>
  <div style="text-align:center;font-size:11px;color:#8497a8;padding:4px 0 10px">能耗分析系统 ${esc((S.boot || {}).version || '')} · © 开心团团</div>`;
}

/* ---------- 返回键：先关弹层，再回总览，再确认退出 ---------- */
history.pushState({ m: 1 }, '');
window.addEventListener('popstate', () => {
  history.pushState({ m: 1 }, '');
  if ($('#modal').style.display === 'flex') return closeModal();
  if ($('#sheet').style.display === 'flex') return closeSheet();
  if (confirmLeave()) return;
  if (S.page !== 'dash') return go('dash');
  if (confirm('退出应用？')) { try { post('auth_logout'); } catch (e) {} location.reload(); }
});
window.addEventListener('beforeunload', e => { if (S.rdDirty) { e.preventDefault(); e.returnValue = ''; } });
window.addEventListener('resize', () => { if (S.page === 'dash' && S.dash) render(); });

boot();
