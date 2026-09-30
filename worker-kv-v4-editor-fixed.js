// @ts-nocheck
/**
 * Cloudflare Worker V4 — KV-first architecture.
 * Required bindings/secrets:
 *   KV namespace binding: APP_CACHE
 *   Secrets/vars: SHEET_ID, CLIENT_EMAIL, PRIVATE_KEY, PASSWORD_PEPPER
 * Cron trigger: 59 * * * *  (hourly at minute 59; includes exactly 16:59 UTC = 23:59 VN)
 */
const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};
const DELETE_CAMPAIGN_ROW_AFTER_PURGE = false; // safer default: retain campaign audit row

addEventListener('fetch', event => event.respondWith(handleRequest(event.request)));
addEventListener('scheduled', event => event.waitUntil(cleanupExpiredCampaigns()));

async function handleRequest(request) {
  if (request.method === 'OPTIONS') return new Response(null, { headers: CORS });
  const url = new URL(request.url);
  try {
    if (request.method === 'GET' && url.pathname === '/api/config') {
      return json(await kvGet('system_config', {}));
    }
    if (request.method === 'POST') {
      const data = await request.json();
      return json(await routeActions(data));
    }
    return new Response('Not Found', { status: 404, headers: CORS });
  } catch (error) {
    return json({ status: 'error', message: error?.message || String(error) }, 500);
  }
}

async function routeActions(data) {
  switch (data.action) {
    // ---------- KV-only reads ----------
    case 'login': return loginFromKv(data);
    case 'get_active_dots': return { status: 'success', dots: await kvGet('active_dots', []) };
    case 'get_canbo_dots': return getCanboDotsFromKv(data.username);
    case 'preview_dot':
    case 'download_dot': return getDotDataFromKv(data.ID_Dot);

    // ---------- Mutations: write Sheets, then mirror KV ----------
    case 'register': return registerAccount(data);
    case 'create_dot': return createDot(data);
    case 'submit_phieuA': return submitPhieuA(data);
    case 'close_dot': return updateDot(data.ID_Dot, { TrangThai: 'Closed' });
    case 'edit_dot': return updateDot(data.ID_Dot, { Ngayketthuc: data.new_date });
    case 'convert_pdf': return convertPdf(data);
    default: return { status: 'fail', message: 'Action không hợp lệ.' };
  }
}

async function loginFromKv(data) {
  const accounts = await kvGet('accounts', {});
  const record = accounts[String(data.username || '')];
  if (!record) return { status: 'fail' };
  const inputHash = await passwordHash(data.password || '');
  return timingSafeEqual(record.password_hash || record, inputHash)
    ? { status: 'success' }
    : { status: 'fail' };
}

async function registerAccount(data) {
  const username = String(data.username || '').trim();
  if (!username || !data.password) return { status: 'fail', message: 'Thiếu tên đăng nhập hoặc mật khẩu.' };
  const accounts = await kvGet('accounts', {});
  if (accounts[username]) return { status: 'fail', message: 'Tên đăng nhập này đã tồn tại, vui lòng chọn tên khác!' };

  const hash = await passwordHash(data.password);
  const token = await getGoogleAuthToken(['https://www.googleapis.com/auth/spreadsheets']);
  await appendRow(token, 'TK_Canbo', [username, `sha256:${hash}`]);
  accounts[username] = { password_hash: hash, created_at: new Date().toISOString() };
  await APP_CACHE.put('accounts', JSON.stringify(accounts));
  return { status: 'success', message: 'Đăng ký thành công!' };
}

async function getCanboDotsFromKv(username) {
  const dots = await kvGet('dots_all', []);
  const mine = dots.filter(d => d.username === username).map(d => ({
    id_dot: d.id_dot,
    ten_dot: `${d.thon} - ${d.phuong} - ${d.tinh}`,
    ngay: formatDate(d.ngay_bat_dau, true),
    ngay_ket_thuc: formatDate(d.ngay_ket_thuc, false),
    trang_thai: d.trang_thai,
    tinh: d.tinh, phuong: d.phuong, thon: d.thon,
  }));
  return { status: 'success', dots: mine };
}

async function getDotDataFromKv(idDot) {
  const dots = await kvGet('dots_all', []);
  const dot = dots.find(d => d.id_dot === idDot);
  if (!dot) return { status: 'fail', message: 'Không tìm thấy đợt rà soát.' };
  const rows = await kvGet(`phieua:${idDot}`, []);
  const Hogiadinh = rows.map((row, index) => ({
    A: index + 1, B: row.B || '', C: row.C ? formatDate(row.C, true) : '',
    '0': row['0'] || '', '1': row['1'] || '', '2': row['2'] || '', '3': row['3'] || '',
    '4': row['4'] || '', '5': row['5'] || '', '6': row['6'] || '', '7': row['7'] || '',
    '8': row['8'] || '', '9': row['9'] || '', D: row.D || '', E: row.E || '',
  }));
  return { status: 'success', Tinh_thanh: dot.tinh, Phuong_xa: dot.phuong, Thon_to: dot.thon, Hogiadinh };
}

async function createDot(data) {
  const token = await getGoogleAuthToken(['https://www.googleapis.com/auth/spreadsheets']);
  const rows = await getSheetValues(token, 'Dot_Rasoat!A1:Z1');
  const h = rows[0] || [];
  const idDot = `DRS_${Date.now()}`;
  const newRow = new Array(h.length).fill('');
  const set = (name, value) => { const i = h.indexOf(name); if (i !== -1) newRow[i] = value; };
  set('ID_Dot', idDot); set('Username', data.username); set('Tinh_thanh', data.tinh);
  set('Phuong_xa', data.phuong); set('Thon_to', data.thon); set('Ngaybatdau', data.ngay_bat_dau);
  set('Ngayketthuc', data.ngay_ket_thuc); set('TrangThai', 'Open');
  set('Dieukien_Bosung', data.dieukien_bosung || '');
  await appendRow(token, 'Dot_Rasoat', newRow);

  const dot = {
    id_dot: idDot, username: data.username, tinh: data.tinh, phuong: data.phuong,
    thon: data.thon, ngay_bat_dau: data.ngay_bat_dau, ngay_ket_thuc: data.ngay_ket_thuc,
    trang_thai: 'Open', dieukien_bosung: data.dieukien_bosung || '', created_at: new Date().toISOString(),
  };
  const dots = await kvGet('dots_all', []); dots.push(dot);
  const active = await kvGet('active_dots', []); active.push(toActiveDot(dot));
  await Promise.all([
    APP_CACHE.put('dots_all', JSON.stringify(dots)),
    APP_CACHE.put('active_dots', JSON.stringify(active)),
  ]);
  return { status: 'success', id_dot: idDot, dot: toActiveDot(dot) };
}

async function submitPhieuA(data) {
  const token = await getGoogleAuthToken(['https://www.googleapis.com/auth/spreadsheets']);
  const rows = await getSheetValues(token, 'Dulieu_Nguoidan!A1:Z1');
  const h = rows[0] || [];
  const newRow = new Array(h.length).fill('');
  const rowObject = {};
  const vnDate = new Intl.DateTimeFormat('vi-VN', { timeZone: 'Asia/Ho_Chi_Minh', day: '2-digit', month: '2-digit', year: 'numeric' }).format(new Date());
  h.forEach((name, i) => {
    let value = name === 'C' ? vnDate : (Object.prototype.hasOwnProperty.call(data, name) ? String(data[name]).replace(/^'/, '') : '');
    newRow[i] = value; rowObject[name] = value;
  });
  await appendRow(token, 'Dulieu_Nguoidan', newRow);
  if (data.ID_Dot) {
    const key = `phieua:${data.ID_Dot}`;
    const list = await kvGet(key, []); list.push(rowObject);
    await APP_CACHE.put(key, JSON.stringify(list));
  }
  return { status: 'success' };
}

async function updateDot(idDot, patch) {
  const dots = await kvGet('dots_all', []);
  const index = dots.findIndex(d => d.id_dot === idDot);
  if (index < 0) return { status: 'fail', message: 'Không tìm thấy đợt rà soát.' };
  const token = await getGoogleAuthToken(['https://www.googleapis.com/auth/spreadsheets']);
  const rows = await getSheetValues(token, 'Dot_Rasoat!A:Z');
  const h = rows[0] || [];
  const rowIndex = rows.findIndex((r, i) => i > 0 && r[h.indexOf('ID_Dot')] === idDot);
  if (rowIndex < 1) return { status: 'fail', message: 'Không tìm thấy đợt trong Sheet.' };

  if (patch.TrangThai !== undefined) await updateCell(token, 'Dot_Rasoat', rowIndex + 1, h.indexOf('TrangThai') + 1, patch.TrangThai);
  if (patch.Ngayketthuc !== undefined) await updateCell(token, 'Dot_Rasoat', rowIndex + 1, h.indexOf('Ngayketthuc') + 1, patch.Ngayketthuc);

  if (patch.TrangThai !== undefined) dots[index].trang_thai = patch.TrangThai;
  if (patch.Ngayketthuc !== undefined) dots[index].ngay_ket_thuc = patch.Ngayketthuc;
  const active = dots.filter(isActiveNow).map(toActiveDot);
  await Promise.all([
    APP_CACHE.put('dots_all', JSON.stringify(dots)),
    APP_CACHE.put('active_dots', JSON.stringify(active)),
  ]);
  return { status: 'success' };
}

async function cleanupExpiredCampaigns() {
  const dots = await kvGet('dots_all', []);
  const purged = new Set(await kvGet('purged_dots', []));
  const expired = dots.filter(d => !purged.has(d.id_dot) && Date.now() >= cleanupAtUtc(d.ngay_ket_thuc));
  if (!expired.length) return;
  const token = await getGoogleAuthToken(['https://www.googleapis.com/auth/spreadsheets']);
  for (const dot of expired) {
    await deleteRowsByIdDot(token, 'Dulieu_Nguoidan', dot.id_dot);
    if (DELETE_CAMPAIGN_ROW_AFTER_PURGE) await deleteRowsByIdDot(token, 'Dot_Rasoat', dot.id_dot);
    else await setDotClosedInSheet(token, dot.id_dot);
    await APP_CACHE.delete(`phieua:${dot.id_dot}`);
    dot.trang_thai = 'Closed'; dot.purged_at = new Date().toISOString(); purged.add(dot.id_dot);
  }
  await Promise.all([
    APP_CACHE.put('dots_all', JSON.stringify(DELETE_CAMPAIGN_ROW_AFTER_PURGE ? dots.filter(d => !purged.has(d.id_dot)) : dots)),
    APP_CACHE.put('active_dots', JSON.stringify(dots.filter(isActiveNow).map(toActiveDot))),
    APP_CACHE.put('purged_dots', JSON.stringify([...purged])),
  ]);
}

async function deleteRowsByIdDot(token, sheetName, idDot) {
  const rows = await getSheetValues(token, `${sheetName}!A:Z`);
  if (!rows.length) return;
  const h = rows[0], idCol = h.indexOf('ID_Dot');
  if (idCol < 0) return;
  const indexes = [];
  rows.forEach((r, i) => { if (i > 0 && r[idCol] === idDot) indexes.push(i); });
  if (!indexes.length) return;
  const sheetId = await getNumericSheetId(token, sheetName);
  const requests = indexes.sort((a,b) => b-a).map(i => ({ deleteDimension: { range: { sheetId, dimension: 'ROWS', startIndex: i, endIndex: i + 1 } } }));
  await sheetsBatchUpdate(token, requests);
}

async function setDotClosedInSheet(token, idDot) {
  const rows = await getSheetValues(token, 'Dot_Rasoat!A:Z');
  if (!rows.length) return;
  const h = rows[0], idCol = h.indexOf('ID_Dot'), statusCol = h.indexOf('TrangThai');
  const i = rows.findIndex((r, x) => x > 0 && r[idCol] === idDot);
  if (i > 0 && statusCol >= 0) await updateCell(token, 'Dot_Rasoat', i + 1, statusCol + 1, 'Closed');
}

async function convertPdf(data) {
  const token = await getGoogleAuthToken(['https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive']);
  const boundary = '-------314159265358979323846';
  const metadata = { name: 'temp_convert.docx', mimeType: 'application/vnd.google-apps.document' };
  const body = `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(metadata)}\r\n--${boundary}\r\nContent-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document\r\nContent-Transfer-Encoding: base64\r\n\r\n${data.docxBase64}\r\n--${boundary}--\r\n`;
  const up = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart', { method:'POST', headers:{ Authorization:`Bearer ${token}`, 'Content-Type':`multipart/related; boundary=${boundary}` }, body });
  const file = await up.json(); if (!file.id) return { status:'error', message:'Upload Google Drive thất bại!' };
  const ex = await fetch(`https://www.googleapis.com/drive/v3/files/${file.id}/export?mimeType=application/pdf`, { headers:{ Authorization:`Bearer ${token}` } });
  const bytes = new Uint8Array(await ex.arrayBuffer()); let binary=''; for (const b of bytes) binary += String.fromCharCode(b);
  await fetch(`https://www.googleapis.com/drive/v3/files/${file.id}`, { method:'DELETE', headers:{ Authorization:`Bearer ${token}` } });
  return { status:'success', pdfBase64:btoa(binary) };
}

function toActiveDot(d) { return { id_dot:d.id_dot, tinh:d.tinh, phuong:d.phuong, thon:d.thon, dieukien_bosung:d.dieukien_bosung || '' }; }
function isActiveNow(d) { return d.trang_thai === 'Open' && Date.now() < cleanupAtUtc(d.ngay_ket_thuc); }
function cleanupAtUtc(value) {
  const m = String(value || '').match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  let y, mo, day;
  if (m) { day=+m[1]; mo=+m[2]; y=+m[3]; }
  else { const d=new Date(value); if (isNaN(d)) return Number.MAX_SAFE_INTEGER; y=d.getUTCFullYear(); mo=d.getUTCMonth()+1; day=d.getUTCDate(); }
  return Date.UTC(y, mo - 1, day + 1, 16, 59, 0); // 23:59 Asia/Ho_Chi_Minh next day
}
function formatDate(value, vn) {
  if (!value) return '';
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(value)) return value;
  const d=new Date(value); if (isNaN(d)) return value;
  const dd=String(d.getDate()).padStart(2,'0'), mm=String(d.getMonth()+1).padStart(2,'0'), yy=d.getFullYear();
  return vn ? `${dd}/${mm}/${yy}` : `${yy}-${mm}-${dd}`;
}
function columnLetter(n) { let s=''; while(n>0){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26);} return s; }
function json(data, status=200) { return new Response(JSON.stringify(data), { status, headers:{...CORS,'Content-Type':'application/json','Cache-Control':'no-store'} }); }
async function kvGet(key, fallback) { const v=await APP_CACHE.get(key,'json'); return v ?? fallback; }
async function passwordHash(password) { const bytes=new TextEncoder().encode(`${password}:${PASSWORD_PEPPER}`); const digest=await crypto.subtle.digest('SHA-256',bytes); return [...new Uint8Array(digest)].map(x=>x.toString(16).padStart(2,'0')).join(''); }
function timingSafeEqual(a,b) { if(a.length!==b.length)return false; let x=0; for(let i=0;i<a.length;i++)x|=a.charCodeAt(i)^b.charCodeAt(i); return x===0; }

async function getSheetValues(token, range) { const r=await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${encodeURIComponent(range)}`,{headers:{Authorization:`Bearer ${token}`}}); const j=await r.json(); if(!r.ok)throw new Error(j.error?.message||'Google Sheets read failed'); return j.values||[]; }
async function appendRow(token, sheet, values) { const r=await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${encodeURIComponent(sheet)}:append?valueInputOption=USER_ENTERED`,{method:'POST',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},body:JSON.stringify({values:[values]})}); if(!r.ok)throw new Error('Google Sheets append failed'); }
async function updateCell(token, sheet, row, col, value) { if(col<1)return; const range=`${sheet}!${columnLetter(col)}${row}`; const r=await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`,{method:'PUT',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},body:JSON.stringify({values:[[value]]})}); if(!r.ok)throw new Error('Google Sheets update failed'); }
async function getNumericSheetId(token,name){const r=await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}?fields=sheets.properties`,{headers:{Authorization:`Bearer ${token}`}});const j=await r.json();const s=(j.sheets||[]).find(x=>x.properties.title===name);if(!s)throw new Error(`Không tìm thấy sheet ${name}`);return s.properties.sheetId;}
async function sheetsBatchUpdate(token,requests){const r=await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}:batchUpdate`,{method:'POST',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},body:JSON.stringify({requests})});if(!r.ok)throw new Error('Google Sheets batchUpdate failed');}
function str2ab(str){return new TextEncoder().encode(str);}
function base64url(source){let binary='';for(const b of new Uint8Array(source))binary+=String.fromCharCode(b);return btoa(binary).replace(/=/g,'').replace(/\+/g,'-').replace(/\//g,'_');}
async function getGoogleAuthToken(scopes){const header=base64url(str2ab(JSON.stringify({alg:'RS256',typ:'JWT'})));const now=Math.floor(Date.now()/1000);const payload=base64url(str2ab(JSON.stringify({iss:CLIENT_EMAIL,scope:scopes.join(' '),aud:'https://oauth2.googleapis.com/token',exp:now+3600,iat:now})));const data=`${header}.${payload}`;const pem=PRIVATE_KEY.replace('-----BEGIN PRIVATE KEY-----','').replace('-----END PRIVATE KEY-----','').replace(/\s/g,'');const raw=Uint8Array.from(atob(pem),c=>c.charCodeAt(0));const key=await crypto.subtle.importKey('pkcs8',raw,{name:'RSASSA-PKCS1-v1_5',hash:'SHA-256'},false,['sign']);const sig=await crypto.subtle.sign('RSASSA-PKCS1-v1_5',key,str2ab(data));const jwt=`${data}.${base64url(sig)}`;const r=await fetch('https://oauth2.googleapis.com/token',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:`grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=${jwt}`});const j=await r.json();if(!j.access_token)throw new Error(j.error_description||'Không lấy được Google token');return j.access_token;}
