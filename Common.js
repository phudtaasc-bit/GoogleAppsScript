const FS_CFG = Object.freeze({
  SHEETS: {
    INPUT: '01. Đầu vào',
    TECH: '01A. Kỹ thuật',
    TECH_LEGACY: '01. Kỹ thuật',
    REVENUE: '02. Doanh thu',
    COST: '03. Chi phí & Vốn',
    PROFIT: '03A. Lợi nhuận & Thuế',
    CASH: '04. Dòng tiền & Tài trợ',
    CASH_LEGACY: '04. Dòng tiền & Lợi nhuận',
    SUMMARY_CASH: '04A. Tổng hợp dòng tiền',
    SUMMARY_CASH_LEGACY: '04A. TH dòng tiền',
    SUMMARY: '00. Tổng hợp',
    CHECKS: '99. Kiểm tra mô hình',
    CHECKS_LEGACY: '99. Checks'
  },
  CONVERGENCE: { tolerance: 1, maxIterations: 100 }
});

function FS_capNhatKyThuat() {
  FS_taoKyThuatTuDauVao();
  const ss = SpreadsheetApp.getActive();
  const legacy = ss.getSheetByName(FS_CFG.SHEETS.TECH_LEGACY);
  let current = ss.getSheetByName(FS_CFG.SHEETS.TECH);
  if (legacy && !current) {
    legacy.setName(FS_CFG.SHEETS.TECH);
    current = legacy;
  }
  current = current || ss.getSheetByName(FS_CFG.SHEETS.TECH);
  if (current) FS_dongBoThongTinSanPhamBoSung_(ss, current);
}

function FS_dongBoThongTinSanPhamBoSung_(ss, tech) {
  const input = ss.getSheetByName(FS_CFG.SHEETS.INPUT);
  if (!input || typeof FS_getTable_ !== 'function' || typeof FS_getByHeaderAny_ !== 'function') return;

  const table = FS_getTable_(input, 'D. CHI TIẾT SẢN PHẨM', 'Loại sản phẩm');
  const titleRow = FS_findExactRow_(tech, 'SAN_PHAM');
  if (!table || !titleRow) return;

  const headerRow = titleRow + 1;
  const sourceByName = {};
  table.rows.forEach(row => {
    const name = FS_getByHeaderAny_(row, table.headers, ['Loại sản phẩm', 'Sản phẩm']);
    if (!name) return;
    sourceByName[FS_key_(name)] = {
      maintenance: FS_getByHeaderAny_(row, table.headers, [
        'Chi phí bảo trì', 'CP bảo trì', 'Bảo trì'
      ]),
      leaseYears: FS_getByHeaderAny_(row, table.headers, [
        'Thời gian thuê (năm)', 'Thời gian thuê', 'Số năm thuê'
      ]),
      landArea: FS_getByHeaderAny_(row, table.headers, [
        'Diện tích đất', 'DT đất'
      ]),
      note: FS_getByHeaderAny_(row, table.headers, ['Ghi chú'])
    };
  });

  tech.getRange(headerRow, 11, 1, 4).setValues([[
    'Chi phí bảo trì', 'Thời gian thuê (năm)', 'Diện tích đất', 'Ghi chú'
  ]]);

  let row = headerRow + 1;
  while (row <= tech.getLastRow()) {
    const productName = String(tech.getRange(row, 1).getDisplayValue() || '').trim();
    if (!productName || /^[A-Z_]+$/.test(productName)) break;
    const src = sourceByName[FS_key_(productName)] || {};
    tech.getRange(row, 11, 1, 4).setValues([[
      src.maintenance || '', src.leaseYears || '', src.landArea || '', src.note || ''
    ]]);
    row++;
  }

  if (row > headerRow + 1) {
    tech.getRange(headerRow + 1, 12, row - headerRow - 1, 1).setNumberFormat('0');
  }
}

function FS_getOrCreateSheet_(ss, name, legacyName) {
  let sh = ss.getSheetByName(name);
  if (sh) return sh;
  if (legacyName) {
    sh = ss.getSheetByName(legacyName);
    if (sh) {
      sh.setName(name);
      return sh;
    }
  }
  return ss.insertSheet(name);
}

function FS_getSheet_(ss, name, legacyName) {
  return ss.getSheetByName(name) || (legacyName ? ss.getSheetByName(legacyName) : null);
}

function FS_resetSheet_(sh, rows, cols) {
  if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
  if (sh.getMaxColumns() < cols) sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
  sh.showRows(1, sh.getMaxRows());
  sh.showColumns(1, sh.getMaxColumns());
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clearContents();
  sh.clearFormats();
}

function FS_norm_(v) {
  return String(v ?? '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd').replace(/Đ/g, 'D')
    .toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}
function FS_key_(v) { return FS_norm_(v).replace(/\s+/g, ''); }
function FS_num_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  const s = String(v ?? '').trim();
  if (!s) return 0;
  const normalized = s.includes(',') && s.includes('.')
    ? s.replace(/\./g, '').replace(',', '.')
    : s.replace(/\s/g, '').replace(',', '.');
  const n = Number(normalized);
  return isFinite(n) ? n : 0;
}
function FS_rate_(v) {
  if (typeof v === 'number') return Math.abs(v) > 1 ? v / 100 : v;
  const s = String(v ?? '').trim();
  const n = Number(s.replace('%', '').replace(',', '.'));
  return isFinite(n) ? (s.includes('%') || Math.abs(n) > 1 ? n / 100 : n) : 0;
}
function FS_rateFromText_(v) {
  if (typeof v === 'number') return FS_rate_(v);
  const m = String(v ?? '').match(/(-?\d+(?:[.,]\d+)?)\s*%/);
  return m ? Number(m[1].replace(',', '.')) / 100 : FS_rate_(v);
}
function FS_namedRateFromText_(v, labels) {
  if (typeof v === 'number') return FS_rate_(v);
  const text = FS_norm_(v);
  for (const label of labels) {
    const k = FS_norm_(label);
    const pos = text.indexOf(k);
    if (pos >= 0) {
      const tail = String(v).slice(pos);
      const m = tail.match(/(-?\d+(?:[.,]\d+)?)\s*%/);
      if (m) return Number(m[1].replace(',', '.')) / 100;
    }
  }
  return 0;
}
function FS_addMonths_(d, m) { const x = new Date(d); return new Date(x.getFullYear(), x.getMonth() + m, x.getDate()); }
function FS_sum_(values) { return values.reduce((s, v) => s + FS_num_(v), 0); }
function FS_ratio_(a, b) { return b ? FS_num_(a) / FS_num_(b) : 0; }

function FS_findExactRow_(sh, text) {
  const target = FS_norm_(text);
  const vals = sh.getRange(1, 1, sh.getLastRow(), 1).getDisplayValues();
  for (let i = 0; i < vals.length; i++) if (FS_norm_(vals[i][0]) === target) return i + 1;
  return 0;
}
function FS_readInfo_(tech) {
  const r = FS_findExactRow_(tech, 'THONG_TIN_CHUNG');
  if (!r) throw new Error('Không tìm thấy THONG_TIN_CHUNG.');
  const o = {};
  for (let i = r + 1; i <= tech.getLastRow(); i++) {
    const k = String(tech.getRange(i, 1).getDisplayValue() || '').trim();
    if (!k || /^[A-Z_]+$/.test(k)) break;
    o[k] = tech.getRange(i, 2).getValue();
  }
  return o;
}
function FS_readBlock_(tech, name) {
  const r = FS_findExactRow_(tech, name);
  if (!r) throw new Error('Không tìm thấy block ' + name);
  const headerRow = r + 1;
  const displayHeaders = tech.getRange(headerRow, 1, 1, tech.getLastColumn()).getDisplayValues()[0];
  const lastCol = Math.max(1, displayHeaders.reduce((n, v, i) => String(v).trim() ? i + 1 : n, 0));
  const rows = [];
  for (let i = r + 2, blank = 0; i <= tech.getLastRow(); i++) {
    const first = String(tech.getRange(i, 1).getDisplayValue() || '').trim();
    if (/^[A-Z_]+$/.test(first)) break;
    const row = tech.getRange(i, 1, 1,lastCol).getValues()[0];
    if (!row.some(v => String(v ?? '').trim() !== '')) { if (++blank >= 3) break; continue; }
    blank = 0;
    rows.push(row);
  }
  return rows;
}
function FS_costMap_(tech) {
  const m = {};
  FS_readBlock_(tech, 'CHI_PHI_CHUNG').forEach(r => {
    const n = String(r[0] || '').trim();
    if (n) m[FS_key_(n)] = { name: n, before: FS_num_(r[1]), vat: FS_rate_(r[2]), after: FS_num_(r[3]), note: String(r[4] || ''), rate: FS_rate_(r[5]) };
  });
  return m;
}
function FS_costItem_(m, names) {
  for (const n of names) { const x = m[FS_key_(n)]; if (x) return x; }
  return { before: 0, vat: 0, rate: 0, name: names[0] };
}

function FS_loaiHinh_(method) { return FS_norm_(method).includes('cho thue') ? 'Cho thuê' : 'Bán'; }
function FS_productType_(name, method) { return FS_loaiHinh_(method) === 'Bán' ? 'SALE' : 'RENT'; }
function FS_maSanPhamGoc_(name) {
  const k = FS_key_(name);
  if (k.includes('noxh') || k.includes('nhaoxahoi')) return 'NOXH';
  if (k.includes('chungcu') || k.includes('canho')) return 'CC';
  if (k.includes('lienke')) return 'LK';
  if (k.includes('shophouse')) return 'SH';
  if (k.includes('tmdv') || k.includes('thuongmaidichvu')) return 'TMDV';
  if (k.includes('cho')) return 'CHO';
  const ascii = FS_norm_(name).split(' ').filter(Boolean).map(x => x[0]).join('').toUpperCase();
  return ascii || 'SP';
}
function FS_ganMaSanPham_(products) {
  const used = {};
  return products.map(p => {
    const base = FS_maSanPhamGoc_(p.name);
    used[base] = (used[base] || 0) + 1;
    return Object.assign({}, p, { code: used[base] === 1 ? base : base + used[base] });
  });
}
function FS_factKey_(monthNo, productCode) { return String(monthNo) + '|' + String(productCode); }
