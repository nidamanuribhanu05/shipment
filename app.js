const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => [...document.querySelectorAll(selector)];
const state = { mode: 'demo', role: 'admin', user: null, csrf: null, records: [], selectedFile: null, preview: null, view: 'overview', busy: false, authenticated: false, limited: false };
const icons = { arrow: '<svg aria-hidden="true"><use href="#i-arrow-up"/></svg>' };
function esc(value) { return String(value ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]); }
function demoRecords() {
  const today = Date.now();
  return [
    { id: 'DEMO-1048', createdAt: new Date(today - 1.8e6).toISOString(), sender: 'Jordan Davis', senderEmail: 'jordan@demo.example', createdBy: 'demo-admin', recipients: ['maya@global.example', 'ops@eu.example'], fileName: 'Export_Manifest_Oct.xlsx', subject: 'October export manifest', rowCount: 42, status: 'submitted', sharePointUrl: null },
    { id: 'DEMO-1047', createdAt: new Date(today - 8.64e7).toISOString(), sender: 'Jordan Davis', senderEmail: 'jordan@demo.example', createdBy: 'demo-admin', recipients: ['supply@apac.example'], fileName: 'APAC_Shipments.xlsx', subject: 'APAC shipment updates', rowCount: 81, status: 'submitted', sharePointUrl: null },
    { id: 'DEMO-1046', createdAt: new Date(today - 1.728e8).toISOString(), sender: 'Alex Morgan', senderEmail: 'alex@demo.example', createdBy: 'demo-other', recipients: ['dispatch@uk.example'], fileName: 'Weekly_Dispatch.xlsx', subject: 'Weekly dispatch', rowCount: 18, status: 'submitted', sharePointUrl: null },
    { id: 'DEMO-1045', createdAt: new Date(today - 2.592e8).toISOString(), sender: 'Jordan Davis', senderEmail: 'jordan@demo.example', createdBy: 'demo-admin', recipients: ['sales@north.example'], fileName: 'Carrier_Tracking.xlsx', subject: 'Carrier tracking file', rowCount: 64, status: 'archived', sharePointUrl: null },
    { id: 'DEMO-1044', createdAt: new Date(today - 3.456e8).toISOString(), sender: 'Alex Morgan', senderEmail: 'alex@demo.example', createdBy: 'demo-other', recipients: ['freight@us.example'], fileName: 'Freight_Orders.xlsx', subject: 'Freight orders', rowCount: 12, status: 'submitted', sharePointUrl: null },
    { id: 'DEMO-1043', createdAt: new Date(today - 4.32e8).toISOString(), sender: 'Jordan Davis', senderEmail: 'jordan@demo.example', createdBy: 'demo-admin', recipients: ['orders@sea.example'], fileName: 'Order_List.xlsx', subject: 'Order list', rowCount: 25, status: 'failed', sharePointUrl: null },
    { id: 'DEMO-1042', createdAt: new Date(today - 5.184e8).toISOString(), sender: 'Alex Morgan', senderEmail: 'alex@demo.example', createdBy: 'demo-other', recipients: ['inventory@eu.example'], fileName: 'Inventory_Export.xlsx', subject: 'Inventory export', rowCount: 110, status: 'submitted', sharePointUrl: null }
  ];
}
function localRecords() {
  try { const stored = JSON.parse(localStorage.getItem('shipflow-demo') || 'null'); return Array.isArray(stored) ? stored : demoRecords(); }
  catch { return demoRecords(); }
}
function saveLocalRecords() { try { localStorage.setItem('shipflow-demo', JSON.stringify(state.records)); } catch {} }
function notify(message, error = false) {
  const el = document.createElement('div');
  el.className = `toast ${error ? 'error' : ''}`;
  el.textContent = message;
  $('#toastContainer').append(el);
  setTimeout(() => el.remove(), 4500);
}
function dateLabel(d) {
  const date = new Date(d);
  return Number.isNaN(date.getTime()) ? 'Unknown' : date.toLocaleDateString(undefined, { day: 'numeric', month: 'short', year: 'numeric' });
}
function cleanStatus(raw) { return ['submitted', 'failed', 'archived'].includes(raw) ? raw : 'archived'; }
function statusHtml(status) { const s = cleanStatus(status); const label = state.mode === 'demo' ? (s === 'failed' ? 'Demo failed' : 'Simulated') : ({ submitted: 'Submitted', failed: 'Failed', archived: 'Archived' })[s]; return `<span class="status ${s}">${label}</span>`; }
function visibleRecords() {
  return state.records.filter(x => state.role === 'admin' || (state.mode === 'demo' ? x.createdBy === 'demo-other' : x.createdBy === state.user?.id));
}
function fileCell(r) {
  return `<div class="filename"><span class="fileglyph">X</span><div><strong title="${esc(r.fileName)}">${esc(r.fileName || 'Shipment workbook')}</strong><span title="${esc(r.subject)}">${esc(r.subject || '')}</span></div></div>`;
}
function recipientText(r) { const a = Array.isArray(r.recipients) ? r.recipients : []; return a.length ? esc(a[0]) + (a.length > 1 ? ` <span title="${esc(a.slice(1).join(', '))}">+${a.length-1}</span>` : '') : '—'; }
function recentRow(r) { return `<tr><td>${fileCell(r)}</td><td>${recipientText(r)}</td><td>${dateLabel(r.createdAt)}</td><td>${statusHtml(r.status)}</td><td>${r.sharePointUrl && /^https:\/\//.test(r.sharePointUrl) ? `<a class="link-icon" href="${esc(r.sharePointUrl)}" target="_blank" rel="noopener noreferrer" title="Open in SharePoint">${icons.arrow}</a>` : ''}</td></tr>`; }
function allRow(r) { return `<tr><td>${fileCell(r)}</td><td>${esc(r.sender || r.senderEmail || 'Unknown')}</td><td>${recipientText(r)}</td><td>${dateLabel(r.createdAt)}</td><td>${statusHtml(r.status)}</td><td>${r.sharePointUrl && /^https:\/\//.test(r.sharePointUrl) ? `<a class="link-icon" href="${esc(r.sharePointUrl)}" target="_blank" rel="noopener noreferrer" title="Open in SharePoint">${icons.arrow}</a>` : '<span style="color:#c2c8d4">—</span>'}</td></tr>`; }
function updateOverview() {
  const all = visibleRecords();
  $('#statTotal').textContent = all.length;
  $('#statSent').textContent = all.filter(x => x.status === 'submitted').length;
  $('#statDest').textContent = new Set(all.flatMap(x => Array.isArray(x.recipients) ? x.recipients.map(e => e.split('@')[1]).filter(Boolean) : [])).size;
  $('#statArchive').textContent = state.mode === 'demo' ? '0' : all.filter(r => r.sharePointUrl).length;
  $('#sideCount').textContent = all.length;
  $('#recentRows').innerHTML = all.slice().sort((a,b) => new Date(b.createdAt) - new Date(a.createdAt)).slice(0,5).map(recentRow).join('') || '<tr><td colspan="5" class="empty-state">No dispatches yet. Create your first shipment.</td></tr>';
  const days = Array.from({length:7}, (_,i) => { const date = new Date(); date.setHours(0,0,0,0); date.setDate(date.getDate() - 6 + i); return date; });
  const counts = days.map(day => all.filter(r => { const d = new Date(r.createdAt); return d >= day && d < new Date(day.getTime() + 86400000); }).length);
  const max = Math.max(1, ...counts);
  $('#activityChart').innerHTML = days.map((day,i) => `<div class="bar-slot" title="${counts[i]} dispatches"><div class="bar-track"><div class="bar-fill" style="height:${Math.max(2, Math.round(counts[i]/max*116))}px"></div></div><span class="bar-day">${day.toLocaleDateString(undefined,{weekday:'short'})}</span></div>`).join('');
  $('#activityTotal').textContent = `${counts.reduce((a,b)=>a+b,0)} this week`;
}
function updateRecords() {
  const search = $('#recordSearch').value.toLowerCase().trim();
  const status = $('#statusFilter').value;
  const visible = visibleRecords().filter(r => (status === 'all' || r.status === status) && (!search || [r.fileName,r.subject,r.sender,r.senderEmail,...(r.recipients||[])].join(' ').toLowerCase().includes(search))).sort((a,b)=>new Date(b.createdAt)-new Date(a.createdAt));
  $('#allRows').innerHTML = visible.map(allRow).join('') || '<tr><td colspan="6" class="empty-state">No matching records found.</td></tr>';
  $('#recordsFooter').textContent = `Showing ${visible.length} dispatch${visible.length === 1 ? '' : 'es'}${state.limited ? ' · Showing latest records only (SharePoint listing limit)' : ''}`;
}
function updateProfile() {
  const isAdmin = state.role === 'admin';
  const name = state.mode === 'demo' ? (isAdmin ? 'Jordan Davis' : 'Alex Morgan') : state.user?.name || 'Guest';
  const shortName = name.trim().split(/\s+/).slice(0,2).map(x=>x[0]?.toUpperCase() || '').join('');
  $('#userName').textContent = name;
  $('#avatar').textContent = shortName || '??';
  $('#roleName').textContent = isAdmin ? 'Administrator' : 'Team member';
  $('#settingRole').textContent = isAdmin ? 'Administrator' : 'Team member';
  $('#greetingName').textContent = name.split(' ')[0] || 'there';
  const hr = new Date().getHours();
  $('.page-heading h1').firstChild.textContent = `${hr < 12 ? 'Good morning' : hr < 18 ? 'Good afternoon' : 'Good evening'}, `;
  $$('.admin-only').forEach(el => el.hidden = !isAdmin);
  if (!isAdmin && state.view === 'access') navigate('overview');
  updateOverview(); updateRecords();
}
function navigate(view) {
  if (view === 'access' && state.role !== 'admin') return;
  state.view = view;
  $$('.page').forEach(p => p.classList.toggle('visible', p.id === `view-${view}`));
  $$('.nav-item').forEach(el => el.classList.toggle('active', el.dataset.view === view));
  $('#crumb').textContent = ({overview:'Overview',new:'New dispatch',sheet:'Excel workspace',records:'Shipment records',access:'Access control',settings:'Workspace info'})[view];
  $('#sidebar').classList.remove('open');
  if (view === 'records') updateRecords();
  window.scrollTo({top:0,behavior:'instant'});
}
function updateModeUI() {
  const demo = state.mode === 'demo';
  $('#demoNotice').hidden = !demo;
  $('#loginNotice').hidden = demo || state.authenticated;
  $('#modePill').textContent = demo ? '● DEMO MODE' : state.authenticated ? '● MICROSOFT 365 SIGNED IN' : '● SIGN IN REQUIRED';
  $('#modePill').classList.toggle('demo', demo || !state.authenticated);
  ['#sharePointState','#mailState','#authState'].forEach(s => { $(s).textContent = demo ? 'Demo mode' : state.authenticated ? 'Configured' : 'Sign in needed'; $(s).classList.toggle('live', !demo && state.authenticated); });
  $('#signoutPanel').hidden = demo || !state.authenticated;
  $('#folderName').textContent = demo ? 'Not connected in demo' : 'Configured SharePoint folder';
  const storage = $('.storage-mini'); if(storage) { storage.querySelector('span').textContent = demo ? 'Not connected · demo only' : 'Microsoft SharePoint'; storage.querySelector('.green-dot').style.background = demo ? '#e3ad63' : '#41c6a3'; }
  const statLabels = $$('.stat-label'); if(statLabels[1]) statLabels[1].textContent = demo ? 'Simulated dispatches' : 'Submitted to email'; if(statLabels[3]) statLabels[3].textContent = demo ? 'SharePoint uploads' : 'Archived files';
  $('#sendLabel').textContent = demo ? 'Simulate dispatch' : 'Send shipment';
}
async function loadRecords() {
  if (state.mode === 'demo') { updateOverview(); updateRecords(); return; }
  if (!state.authenticated || !state.user?.role) return;
  try {
    const res = await fetch('/api/shipments', {credentials:'same-origin'});
    const json = await res.json();
    if (!res.ok) throw new Error(json.error || 'Unable to load records');
    state.records = json.records || [];
    state.limited = Boolean(json.limited);
    updateOverview(); updateRecords();
  } catch (err) { notify(err.message, true); }
}
async function chooseFile(file) {
  if (!file) return;
  if (!/\.xlsx$/i.test(file.name)) return notify('Choose an Excel .xlsx file.', true);
  if (file.size > 2500000) return notify('File must be 2.5 MB or smaller.', true);
  try {
    const preview = await window.ShipflowWorkbook.load(file);
    state.selectedFile = file;
    state.preview = preview;
    $('#selectedFile').hidden = false;
    $('#selectedFileName').textContent = file.name;
    $('#selectedFileInfo').textContent = `${(file.size / 1024).toFixed(0)} KB · .XLSX workbook · Click Excel workspace to edit`;
    $('#dropzone').hidden = true;
    $('#filePreview').hidden = false;
    $('#previewRows').textContent = `${preview.rowCount} data rows · ${preview.sheetName}`;
    $('#previewHead').innerHTML = `<tr>${preview.headers.map(v => `<th>${esc(v)}</th>`).join('')}</tr>`;
    $('#previewBody').innerHTML = preview.preview.map(row => `<tr>${row.map(v => `<td>${esc(v)}</td>`).join('')}</tr>`).join('');
  } catch(err) { notify(err.message || 'Could not open workbook.', true); }
}
function resetFile() {
  state.selectedFile = null;
  state.preview = null;
  $('#fileInput').value = '';
  $('#selectedFile').hidden = true;
  $('#dropzone').hidden = false;
  $('#filePreview').hidden = true;
}
function parseRecipients(value) {
  const a = [...new Set(value.split(/[;,\n]+/).map(x=>x.trim().toLowerCase()).filter(Boolean))];
  if (!a.length || a.length > 20 || a.some(x=>! /^[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]+$/.test(x))) throw new Error('Enter 1 to 20 valid recipient email addresses.');
  return a;
}
function showModal(record) {
  $('#modalHeading').textContent = state.mode === 'demo' ? 'Demo dispatch created!' : record.status === 'submitted' ? 'Shipment submitted!' : 'Sending failed';
  $('#modalText').textContent = state.mode === 'demo' ? 'This is a local simulation. No email was sent or SharePoint file stored.' : record.status === 'submitted' ? 'Your workbook is archived in SharePoint. Microsoft 365 accepted the email request, but delivery is not yet confirmed.' : `Your Excel file was archived, but the email failed: ${record.error || 'Unknown error'}`;
  $('#modalRef').textContent = `Shipment reference: ${record.id}`;
  $('#successModal').hidden = false;
}
async function sendShipment(e) {
  e.preventDefault();
  if (state.busy) return;
  if (!state.selectedFile) return notify('Please choose a workbook.', true);
  if (state.mode === 'live' && !state.authenticated) return notify('Sign in with Microsoft 365 first.', true);
  if (state.mode === 'live' && !state.user?.role) return notify('Your account needs a Shipment.User or Shipment.Admin role.', true);
  let recipients;
  try { recipients = parseRecipients($('#recipients').value); } catch(e) { return notify(e.message, true); }
  const subject = $('#subject').value.trim();
  const message = $('#message').value.trim();
  if (!subject) return notify('Email subject is required.', true);
  state.busy = true;
  $('#sendButton').disabled = true;
  $('#sendLabel').textContent = state.mode === 'demo' ? 'Simulating...' : 'Archiving and sending...';
  try {
    let record;
    if (state.mode === 'demo') {
      record = { id: `DEMO-${Math.round(Math.random()*899999+100000)}`, createdAt:new Date().toISOString(), createdBy:'demo-admin', sender:'Jordan Davis', senderEmail:'jordan@demo.example', recipients,subject,message,fileName:state.selectedFile.name,fileSize:state.selectedFile.size,rowCount:0,status:'submitted',sharePointUrl:null };
      if (state.role === 'user') { record.sender='Alex Morgan';record.createdBy='demo-other'; }
      state.records.unshift(record);saveLocalRecords();
      await new Promise(resolve=>setTimeout(resolve,520));
    } else {
      const data = new FormData();
      data.append('file', state.selectedFile);data.append('recipients', recipients.join(','));data.append('subject', subject);data.append('message', message);
      const resp = await fetch('/api/send',{method:'POST',headers:{'X-CSRF-Token':state.csrf},body:data,credentials:'same-origin'});
      record = await resp.json();
      if (!resp.ok && !record.id) throw new Error(record.error || 'Shipment could not be submitted.');
      await loadRecords();
    }
    $('#sendForm').reset();resetFile();updateOverview();updateRecords();showModal(record);
  } catch(err) { notify(err.message, true); }
  finally { state.busy=false;$('#sendButton').disabled=false;$('#sendLabel').textContent=state.mode === 'demo' ? 'Simulate dispatch' : 'Send shipment'; }
}
function initEvents() {
  $$('[data-view]').forEach(el=>el.addEventListener('click',()=>navigate(el.dataset.view)));
  $$('[data-go]').forEach(el=>el.addEventListener('click',()=>navigate(el.dataset.go)));
  $('#mobileMenu').addEventListener('click',()=>$('#sidebar').classList.toggle('open'));
  $('#roleSwitch').addEventListener('change',e=>{state.role=e.target.value;updateProfile();});
  $('#profileButton').addEventListener('click',()=>navigate('settings'));
  $('#recordSearch').addEventListener('input',updateRecords);
  $('#statusFilter').addEventListener('change',updateRecords);
  $('#refreshButton').addEventListener('click',()=>{if(state.mode==='demo'){notify('Demo records are stored locally in this browser.');updateOverview();updateRecords();} else loadRecords();});
  $('#dropzone').addEventListener('click',()=>$('#fileInput').click());
  $('#dropzone').addEventListener('keydown',e=>{if(e.key==='Enter' || e.key===' '){e.preventDefault();$('#fileInput').click();}});
  $('#fileInput').addEventListener('change',e=>chooseFile(e.target.files[0]));
  $('#removeFile').addEventListener('click',resetFile);
  $('#dropzone').addEventListener('dragover',e=>{e.preventDefault();$('#dropzone').classList.add('dragging');});
  $('#dropzone').addEventListener('dragleave',()=>$('#dropzone').classList.remove('dragging'));
  $('#dropzone').addEventListener('drop',e=>{e.preventDefault();$('#dropzone').classList.remove('dragging');chooseFile(e.dataTransfer.files[0]);});
  $('#sendForm').addEventListener('submit',sendShipment);
  $('#modalRecords').addEventListener('click',()=>{$('#successModal').hidden=true;navigate('records');});
  $('#modalAgain').addEventListener('click',()=>{$('#successModal').hidden=true;navigate('new');});
  $('#successModal').addEventListener('click',e=>{if(e.target.id==='successModal')$('#successModal').hidden=true;});
  document.addEventListener('keydown',e=>{if(e.key==='Escape'){$('#successModal').hidden=true;$('#sidebar').classList.remove('open');}});
  $('#signoutButton').addEventListener('click',async()=>{try {await fetch('/auth/logout',{method:'POST',headers:{'X-CSRF-Token':state.csrf}}); location.reload();} catch(e){notify('Could not sign out.',true);}});
}
async function boot() {
  window.ShipflowWorkbook.init();
  initEvents();
  let session;
  try { const response = await fetch('/api/session');if(response.ok)session=await response.json(); }
  catch {}
  state.mode = session?.mode === 'live' ? 'live' : 'demo';
  state.authenticated = Boolean(session?.authenticated);
  state.user = session?.user || null;
  state.csrf = session?.csrf || null;
  state.role = state.mode === 'live' ? (session?.user?.role || 'user') : 'admin';
  state.records = state.mode === 'demo' ? localRecords() : [];
  updateModeUI();updateProfile();
  if (state.mode === 'live' && state.authenticated && !state.user?.role) notify('An administrator needs to assign you an Entra app role.', true);
  await loadRecords();
}
window.shipflowNotify = notify;
window.shipflowSelectFile = chooseFile;
window.shipflowNavigate = navigate;
boot();
