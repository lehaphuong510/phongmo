/**
 * GOOGLE APPS SCRIPT HOÀN CHỈNH — KV-FIRST V4
 * 1) Điền 4 giá trị bên dưới.
 * 2) Dùng cùng PASSWORD_PEPPER với secret PASSWORD_PEPPER bên Cloudflare Worker.
 * 3) Nút hiện tại vẫn gán cho hàm syncToCloudflare.
 */
const CF_ACCOUNT_ID = '';
const CF_NAMESPACE_ID = '';
const CF_API_TOKEN = '';
const PASSWORD_PEPPER = ''; // tự đặt chuỗi dài, ví dụ 40+ ký tự; phải giống Worker

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Hệ Thống Rà Soát')
    .addItem('Đồng bộ TOÀN BỘ dữ liệu lên Web', 'syncToCloudflare')
    .addToUi();
}

/** Hàm chính cho menu và nút “Cập nhật web” hiện tại. */
function syncToCloudflare() {
  validateSettings_();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const systemConfig = buildSystemConfig_(ss);
    putKvJson_('system_config', systemConfig);
    syncOperationalDataToKv_(ss);
    console.log('✅ Đã đồng bộ cấu hình, tài khoản, đợt rà soát và Phiếu A lên Web!');
    try {
      SpreadsheetApp.getActive().toast('Đồng bộ Cloudflare KV hoàn tất.', '✅ Thành công', 6);
      SpreadsheetApp.getUi().alert('✅ Đã đồng bộ TOÀN BỘ dữ liệu lên Web thành công!');
    } catch (e) {}
  } catch (e) {
    console.error(e);
    try { SpreadsheetApp.getUi().alert('❌ Lỗi đồng bộ: ' + e.message); } catch (err) {}
    throw e;
  } finally {
    lock.releaseLock();
  }
}

/** Tên tương thích với file hướng dẫn cũ; chạy hàm này cũng giống syncToCloudflare. */
function syncAllToWeb() {
  return syncToCloudflare();
}

/**
 * Chỉ tạo trigger này nếu có Google Form ghi TRỰC TIẾP vào Spreadsheet.
 * Website ghi qua Worker đã tự cập nhật KV nên không cần chạy hàm này.
 */
function syncWhenFormSubmitted(e) {
  validateSettings_();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try { syncOperationalDataToKv_(SpreadsheetApp.getActiveSpreadsheet()); }
  finally { lock.releaseLock(); }
}

// ============================================================
// A. TẠO SYSTEM_CONFIG — GIỮ NGUYÊN LOGIC CODE CŨ
// ============================================================
function buildSystemConfig_(ss) {
  // 1. Địa chính
  const sheetDiaChinh = requiredSheet_(ss, 'Diachinh').getDataRange().getValues();
  let diachinhJSON = {};
  for (let i = 1; i < sheetDiaChinh.length; i++) {
    let tinh = sheetDiaChinh[i][0], phuong = sheetDiaChinh[i][1];
    if (!tinh) continue;
    if (!diachinhJSON[tinh]) diachinhJSON[tinh] = [];
    if (phuong && !diachinhJSON[tinh].includes(phuong)) diachinhJSON[tinh].push(phuong);
  }

  // 2. Phiếu A
  const sheetCauHoi = requiredSheet_(ss, 'PhieuA_Cauhoi').getDataRange().getValues();
  const sheetLuaChon = requiredSheet_(ss, 'PhieuA_Luachon').getDataRange().getValues();
  let cauhoiJSON = {};
  for (let i = 1; i < sheetCauHoi.length; i++) {
    let ma = String(sheetCauHoi[i][0] || '').trim();
    if (ma) cauhoiJSON[ma] = { noi_dung: sheetCauHoi[i][1], lua_chon: [] };
  }
  for (let i = 1; i < sheetLuaChon.length; i++) {
    let ma = String(sheetLuaChon[i][0] || '').trim();
    if (ma && cauhoiJSON[ma]) cauhoiJSON[ma].lua_chon.push(sheetLuaChon[i][1]);
  }

  // 3. Khu vực & Điểm cộng Đô thị
  const sheetDCKV = requiredSheet_(ss, 'DiaChinh_KV').getDataRange().getValues();
  let dictKhuVuc = {};
  let dckvHeaders = sheetDCKV[0].map(h => String(h).trim());
  let idxKinhTeNT = dckvHeaders.indexOf('Khu_vuc_kinh_te_NT');
  for (let i = 1; i < sheetDCKV.length; i++) {
    if (sheetDCKV[i][0]) {
      dictKhuVuc[String(sheetDCKV[i][0]).trim()] = {
        chung: String(sheetDCKV[i][1] || '').trim(),
        nt_rieng: idxKinhTeNT !== -1 ? String(sheetDCKV[i][idxKinhTeNT] || '').trim() : ''
      };
    }
  }

  const sheetDCDT = requiredSheet_(ss, 'PhieuB1_DiemCongDoThi').getDataRange().getValues();
  let dictDiemCong = {};
  let dcdtHeaders = sheetDCDT[0].map(h => String(h).trim());
  let idxKV = dcdtHeaders.indexOf('Khu_vuc_kinh_te');
  let idxMa = dcdtHeaders.indexOf('Ma_Vung');
  let idxDiem = dcdtHeaders.indexOf('Diem_Cong');
  for (let i = 1; i < sheetDCDT.length; i++) {
    if (sheetDCDT[i][idxKV]) {
      dictDiemCong[String(sheetDCDT[i][idxKV]).trim()] = {
        ma: String(sheetDCDT[i][idxMa] || '').trim(),
        diem: Number(sheetDCDT[i][idxDiem]) || 0
      };
    }
  }

  // 4. Dân tộc
  const sheetDantoc = ss.getSheetByName('Ma_Dantoc');
  let listDantoc = [];
  if (sheetDantoc) {
    const dataDantoc = sheetDantoc.getDataRange().getValues();
    let dtHeaders = dataDantoc[0].map(h => String(h).trim());
    let idxMaDT = dtHeaders.indexOf('Ma_Dantoc'), idxTenDT = dtHeaders.indexOf('Dantoc');
    for (let i = 1; i < dataDantoc.length; i++) {
      if (dataDantoc[i][idxMaDT]) {
        listDantoc.push({ ma: String(dataDantoc[i][idxMaDT]), ten: String(dataDantoc[i][idxTenDT]) });
      }
    }
  }

  // 5. Phiếu B1
  const sheetB1 = requiredSheet_(ss, 'PhieuB1_Barem').getDataRange().getValues();
  let b1Headers = sheetB1[0].map(h => String(h).trim());
  let idxMulti = b1Headers.indexOf('Multichoice');
  let b1Questions = [];
  let currentQ = null;
  for (let i = 1; i < sheetB1.length; i++) {
    let maCauHoi = String(sheetB1[i][b1Headers.indexOf('Ma_Cauhoi')] || '').trim();
    let noiDung = String(sheetB1[i][b1Headers.indexOf('Noidung_Cauhoi')] || '').trim();
    let ghiChu = b1Headers.indexOf('Ghi_chu') !== -1 ? String(sheetB1[i][b1Headers.indexOf('Ghi_chu')] || '').trim() : '';
    let luachon = String(sheetB1[i][b1Headers.indexOf('Noidung_Luachon')] || '').trim();
    let isMulti = idxMulti !== -1 ? String(sheetB1[i][idxMulti] || '').trim().toLowerCase() === 'x' : false;
    if (maCauHoi) {
      if (!currentQ || currentQ.ma !== maCauHoi) {
        if (currentQ) b1Questions.push(currentQ);
        currentQ = { ma: maCauHoi, noidung: noiDung, ghichu: ghiChu, isMulti: isMulti, options: [] };
      } else {
        if (noiDung) currentQ.noidung = noiDung;
        if (ghiChu) currentQ.ghichu = ghiChu;
        if (isMulti) currentQ.isMulti = true;
      }
    }
    if (currentQ && luachon) {
      let opt = { text: luachon };
      ['NT_DBSH', 'NT_MNPB', 'NT_BTB', 'NT_DHTN', 'NT_DNB', 'NT_DBSCL', 'DoThi_Chung'].forEach(col => {
        if (b1Headers.indexOf(col) !== -1) opt[col] = Number(sheetB1[i][b1Headers.indexOf(col)]) || 0;
      });
      currentQ.options.push(opt);
    }
  }
  if (currentQ) b1Questions.push(currentQ);

  // 6. Phiếu C
  const sheetC = requiredSheet_(ss, 'PhieuC_Cauhoi').getDataRange().getValues();
  let cHeaders = sheetC[0].map(h => String(h).trim());
  let phieuCQuestions = [];
  let currentC = null;
  for (let i = 1; i < sheetC.length; i++) {
    let row = sheetC[i];
    let ma = String(row[cHeaders.indexOf('Ma_Cauhoi')] || '').trim();
    let noidung = String(row[cHeaders.indexOf('Noidung_Cauhoi')] || '').trim();
    let luachon = String(row[cHeaders.indexOf('Noidung_Luachon')] || '').trim();
    let maLuaChon = String(row[cHeaders.indexOf('Ma_Luachon')] || '').trim();
    let nextQ = cHeaders.indexOf('Next_question') !== -1 ? String(row[cHeaders.indexOf('Next_question')] || '').trim() : '';
    let dieuKien = cHeaders.indexOf('Dieu_kien_tuoi') !== -1 ? String(row[cHeaders.indexOf('Dieu_kien_tuoi')] || '').trim() : '';
    let kBatBuoc = cHeaders.indexOf('Khong_bat_buoc') !== -1 ? String(row[cHeaders.indexOf('Khong_bat_buoc')] || '').trim() : '';
    let gChu = cHeaders.indexOf('Ghi_chu') !== -1 ? String(row[cHeaders.indexOf('Ghi_chu')] || '').trim() : '';
    if (ma) {
      if (!currentC || currentC.ma !== ma) {
        if (currentC) phieuCQuestions.push(currentC);
        currentC = { ma, noidung, ghichu: gChu, dieu_kien: dieuKien, k_bat_buoc: kBatBuoc, isMulti: gChu.includes('Multichoice'), isHo: gChu.includes('Cau_hoi_ho'), options: [] };
      } else {
        if (noidung) currentC.noidung = noidung;
        if (gChu) { currentC.ghichu = gChu; currentC.isMulti = gChu.includes('Multichoice'); currentC.isHo = gChu.includes('Cau_hoi_ho'); }
        if (dieuKien) currentC.dieu_kien = dieuKien;
      }
    }
    if (currentC && luachon) currentC.options.push({ text: luachon, ma_lc: maLuaChon, next: nextQ });
  }
  if (currentC) phieuCQuestions.push(currentC);

  // 7. Phiếu B2
  const sheetB2List = requiredSheet_(ss, 'PhieuB2_Cauhoi').getDataRange().getValues();
  let b2HeadersMap = sheetB2List[0].map(h => String(h).trim());
  let phieuB2Questions = [];
  for (let i = 1; i < sheetB2List.length; i++) {
    let ma = String(sheetB2List[i][b2HeadersMap.indexOf('Ma_Cauhoi')] || '').trim();
    if (ma) phieuB2Questions.push({
      ma,
      section: String(sheetB2List[i][b2HeadersMap.indexOf('Section')] || '').trim(),
      noidung: String(sheetB2List[i][b2HeadersMap.indexOf('Noidung_Cauhoi')] || '').trim()
    });
  }

  // 8. Chuẩn dinh dưỡng
  const sheetDD = requiredSheet_(ss, 'Chuan_DinhDuong').getDataRange().getValues();
  let ddHeaders = sheetDD[0].map(h => String(h).trim());
  let chuanDD = [];
  for (let i = 1; i < sheetDD.length; i++) {
    if (String(sheetDD[i][0] || '').trim()) {
      chuanDD.push({
        gioi_tinh: String(sheetDD[i][ddHeaders.indexOf('Gioi_tinh')] || '').trim(),
        thang_tuoi: Number(sheetDD[i][ddHeaders.indexOf('Thang_tuoi')] || 0),
        can_nang: Number(sheetDD[i][ddHeaders.indexOf('Can_nang_min')] || 0),
        chieu_cao: Number(sheetDD[i][ddHeaders.indexOf('Chieu_cao_min')] || 0)
      });
    }
  }

  return {
    diachinh: diachinhJSON,
    cauhoi: cauhoiJSON,
    diachinh_kv: dictKhuVuc,
    diemcong_dothi: dictDiemCong,
    madantoc: listDantoc,
    phieu_b1: b1Questions,
    phieu_c: phieuCQuestions,
    phieu_b2_list: phieuB2Questions,
    chuan_dinh_duong: chuanDD
  };
}

// ============================================================
// B. ĐỒNG BỘ DỮ LIỆU VẬN HÀNH LÊN KV
// ============================================================
function syncOperationalDataToKv_(ss) {
  // 1. Tài khoản: KV chỉ lưu hash, không đưa mật khẩu gốc lên KV.
  const accountValues = getSheetValues_(ss, 'TK_Canbo');
  let accounts = {};
  for (let i = 1; i < accountValues.length; i++) {
    let username = String(accountValues[i][0] || '').trim();
    let stored = String(accountValues[i][1] || '');
    if (!username || !stored) continue;
    let hash = stored.indexOf('sha256:') === 0 ? stored.substring(7) : sha256Hex_(stored + ':' + PASSWORD_PEPPER);
    accounts[username] = { password_hash: hash };
  }
  putKvJson_('accounts', accounts);

  // 2. Tất cả đợt và đợt đang mở.
  const previousDots = getKvJson_('dots_all', []);
  const previousById = {}; previousDots.forEach(d => previousById[d.id_dot] = d);
  const purgedIds = new Set(getKvJson_('purged_dots', []));
  const dotRows = tableObjects_(ss, 'Dot_Rasoat');
  let dots = dotRows.map(r => ({
    id_dot: String(r.ID_Dot || ''),
    username: String(r.Username || ''),
    tinh: String(r.Tinh_thanh || ''),
    phuong: String(r.Phuong_xa || ''),
    thon: String(r.Thon_to || ''),
    ngay_bat_dau: normalizeDate_(r.Ngaybatdau),
    ngay_ket_thuc: normalizeDate_(r.Ngayketthuc),
    trang_thai: String(r.TrangThai || ''),
    dieukien_bosung: String(r.Dieukien_Bosung || '')
  })).filter(d => d.id_dot);

  let active = dots.filter(isActiveDot_).map(d => ({
    id_dot: d.id_dot,
    tinh: d.tinh,
    phuong: d.phuong,
    thon: d.thon,
    dieukien_bosung: d.dieukien_bosung
  }));
  putKvJson_('active_dots', active);

  // 3. Dữ liệu Phiếu A theo ID_Dot.
  const submissions = tableObjects_(ss, 'Dulieu_Nguoidan');
  let grouped = {};
  submissions.forEach(row => {
    let id = String(row.ID_Dot || '');
    if (!id) return;
    if (!grouped[id]) grouped[id] = [];
    grouped[id].push(row);
  });

  dots.forEach(d => {
    const previous = previousById[d.id_dot] || {};
    if(purgedIds.has(d.id_dot)) {
      d.survey_count = previous.survey_count === undefined ? null : previous.survey_count;
      d.submission_count = previous.submission_count === undefined ? null : previous.submission_count;
      if(previous.purged_at) d.purged_at = previous.purged_at;
    } else {
      const rows = grouped[d.id_dot] || []; const ids = new Set(); let withoutId = 0;
      rows.forEach(row => { const id = String(row.CCCD || '').replace(/^'/, '').trim(); if(id) ids.add(id); else withoutId++; });
      d.survey_count = ids.size + withoutId; d.submission_count = rows.length;
    }
  });
  putKvJson_('dots_all', dots);
  let oldKeys = getKvJson_('phieua_keys', []);
  let newKeys = Object.keys(grouped);
  newKeys.forEach(id => putKvJson_('phieua:' + id, grouped[id]));
  oldKeys.filter(id => !newKeys.includes(id)).forEach(id => deleteKv_('phieua:' + id));
  putKvJson_('phieua_keys', newKeys);
}

// ============================================================
// C. HELPERS
// ============================================================
function validateSettings_() {
  if (!CF_ACCOUNT_ID) throw new Error('Chưa điền CF_ACCOUNT_ID ở đầu file.');
  if (!CF_NAMESPACE_ID) throw new Error('Chưa điền CF_NAMESPACE_ID ở đầu file.');
  if (!CF_API_TOKEN) throw new Error('Chưa điền CF_API_TOKEN ở đầu file.');
  if (!PASSWORD_PEPPER || PASSWORD_PEPPER.length < 24) throw new Error('PASSWORD_PEPPER phải có ít nhất 24 ký tự và giống secret bên Worker.');
}

function requiredSheet_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Không tìm thấy sheet: ' + name);
  return sheet;
}

function getSheetValues_(ss, name) {
  const sheet = requiredSheet_(ss, name);
  if (!sheet.getLastRow() || !sheet.getLastColumn()) return [];
  return sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
}

function tableObjects_(ss, name) {
  const values = getSheetValues_(ss, name);
  if (!values.length) return [];
  const headers = values[0].map(String);
  return values.slice(1)
    .filter(row => row.some(v => v !== ''))
    .map(row => {
      let obj = {};
      headers.forEach((h, i) => { if (h) obj[h] = row[i] instanceof Date ? normalizeDate_(row[i]) : row[i]; });
      return obj;
    });
}

function normalizeDate_(value) {
  if (!value) return '';
  if (value instanceof Date) return Utilities.formatDate(value, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy');
  const text = String(value).trim();
  const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  return iso ? iso[3] + '/' + iso[2] + '/' + iso[1] : text.split(' ')[0];
}

function isActiveDot_(d) {
  if (d.trang_thai !== 'Open') return false;
  const m = String(d.ngay_ket_thuc).match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!m) return true;
  const cutoff = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]) + 1, 23, 59, 0);
  return new Date() < cutoff;
}

function sha256Hex_(text) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8)
    .map(b => { const v = b < 0 ? b + 256 : b; return ('0' + v.toString(16)).slice(-2); })
    .join('');
}

function kvUrl_(key) {
  return `https://api.cloudflare.com/client/v4/accounts/${CF_ACCOUNT_ID}/storage/kv/namespaces/${CF_NAMESPACE_ID}/values/${encodeURIComponent(key)}`;
}

function putKvJson_(key, value) {
  const response = UrlFetchApp.fetch(kvUrl_(key), {
    method: 'put',
    headers: { Authorization: `Bearer ${CF_API_TOKEN}` },
    contentType: 'application/json',
    payload: JSON.stringify(value),
    muteHttpExceptions: true
  });
  assertCloudflareOk_(response, 'PUT ' + key);
}

function getKvJson_(key, fallback) {
  const response = UrlFetchApp.fetch(kvUrl_(key), {
    method: 'get',
    headers: { Authorization: `Bearer ${CF_API_TOKEN}` },
    muteHttpExceptions: true
  });
  if (response.getResponseCode() === 404) return fallback;
  assertCloudflareOk_(response, 'GET ' + key);
  try { return JSON.parse(response.getContentText()); } catch (e) { return fallback; }
}

function deleteKv_(key) {
  const response = UrlFetchApp.fetch(kvUrl_(key), {
    method: 'delete',
    headers: { Authorization: `Bearer ${CF_API_TOKEN}` },
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 404) assertCloudflareOk_(response, 'DELETE ' + key);
}

function assertCloudflareOk_(response, action) {
  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Cloudflare ' + action + ' thất bại — HTTP ' + code + ': ' + response.getContentText());
  }
}
