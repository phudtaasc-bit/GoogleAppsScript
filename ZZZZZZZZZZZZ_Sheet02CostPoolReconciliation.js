/*************************************************
 * ZZZZZZZZZZZZ_Sheet02CostPoolReconciliation.js
 * Đồng bộ tổng giá vốn phân bổ Sheet 02 với nguồn chi tiết.
 *
 * Kiểm soát bắt buộc:
 * - Tổng Tiền SDĐ/thuê đất phân bổ Sheet 02 = nguồn Sheet 03.
 * - Tổng Chi phí dự phòng phân bổ Sheet 02 = nguồn Sheet 03.
 * - Tổng Chi phí lãi vay phân bổ Sheet 02 = Lãi vay vốn hóa Sheet 04.
 * - Sau khi hiệu chỉnh, tính lại Tổng giá vốn, Lợi nhuận chịu thuế, Thuế TNDN.
 *************************************************/

const FS02_POOL_BASE_LAP_SHEET02_ = FS_lapSheet02;

FS_lapSheet02 = function() {
  FS02_POOL_BASE_LAP_SHEET02_();
  FS02_POOL_reconcile_();
};

function FS02_POOL_reconcile_() {
  const ss = SpreadsheetApp.getActive();
  const sh02 = ss.getSheetByName('02. Doanh thu');
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');
  if (!sh02 || !sh03 || !sh04) {
    throw new Error('Thiếu Sheet 02, 03 hoặc 04 để đối chiếu giá vốn phân bổ.');
  }

  const m02 = FS02_POOL_detectHeader_(sh02, [
    'Tổng doanh thu trước VAT', 'Thuế TNDN %',
    'Tiền SDĐ/thuê đất phân bổ trước VAT',
    'Chi phí dự phòng phân bổ trước VAT',
    'Chi phí lãi vay phân bổ', 'Tổng giá vốn tính thuế'
  ]);
  const m03 = FS02_POOL_detectHeader_(sh03, [
    'Tiền SDĐ/thuê đất trước VAT', 'Chi phí dự phòng trước VAT'
  ]);
  const m04 = FS02_POOL_detectHeader_(sh04, ['Lãi vay vốn hóa']);

  const c02 = {
    revenue: FS02_POOL_requireCol_(m02.headers, ['Tổng doanh thu trước VAT']),
    citRate: FS02_POOL_requireCol_(m02.headers, ['Thuế TNDN %']),
    cpxd: FS02_POOL_requireCol_(m02.headers, ['CP XD/TB trực tiếp trước VAT']),
    selling: FS02_POOL_requireCol_(m02.headers, ['Chi phí bán hàng trước VAT']),
    operating: FS02_POOL_requireCol_(m02.headers, ['Chi phí vận hành thuê trước VAT']),
    gpmb: FS02_POOL_requireCol_(m02.headers, ['Chi phí GPMB phân bổ trước VAT']),
    htkt: FS02_POOL_requireCol_(m02.headers, ['Chi phí HTKT phân bổ trước VAT']),
    land: FS02_POOL_requireCol_(m02.headers, ['Tiền SDĐ/thuê đất phân bổ trước VAT']),
    reserve: FS02_POOL_requireCol_(m02.headers, ['Chi phí dự phòng phân bổ trước VAT']),
    interest: FS02_POOL_requireCol_(m02.headers, ['Chi phí lãi vay phân bổ']),
    totalCost: FS02_POOL_requireCol_(m02.headers, ['Tổng giá vốn tính thuế']),
    taxableProfit: FS02_POOL_requireCol_(m02.headers, ['Lợi nhuận chịu thuế']),
    cit: FS02_POOL_requireCol_(m02.headers, ['Thuế TNDN tạm tính'])
  };

  const sourceLand = FS02_POOL_sumNamed_(sh03, m03, 'Tiền SDĐ/thuê đất trước VAT');
  const sourceReserve = FS02_POOL_sumNamed_(sh03, m03, 'Chi phí dự phòng trước VAT');
  const sourceInterest = FS02_POOL_sumNamed_(sh04, m04, 'Lãi vay vốn hóa');

  const n = sh02.getLastRow() - m02.headerRow;
  if (n <= 0) return;

  const width = sh02.getLastColumn();
  const rows = sh02.getRange(m02.headerRow + 1, 1, n, width).getValues();

  FS02_POOL_reconcileColumn_(rows, c02.land - 1, sourceLand, row => {
    return Math.max(0, FS02_POOL_num_(row[c02.land - 1]));
  });

  FS02_POOL_reconcileColumn_(rows, c02.reserve - 1, sourceReserve, row => {
    // Dự phòng gắn với XD/TB + HTKT, đúng nguyên tắc đã chốt.
    return Math.max(0,
      FS02_POOL_num_(row[c02.cpxd - 1]) +
      FS02_POOL_num_(row[c02.htkt - 1])
    );
  });

  FS02_POOL_reconcileColumn_(rows, c02.interest - 1, sourceInterest, row => {
    // Giữ nguyên profile phân bổ lãi vay hiện có; nếu profile bằng 0 thì dùng giá vốn đầu tư làm trọng số.
    const current = Math.max(0, FS02_POOL_num_(row[c02.interest - 1]));
    if (current > 0) return current;
    return Math.max(0,
      FS02_POOL_num_(row[c02.cpxd - 1]) +
      FS02_POOL_num_(row[c02.gpmb - 1]) +
      FS02_POOL_num_(row[c02.htkt - 1]) +
      FS02_POOL_num_(row[c02.land - 1]) +
      FS02_POOL_num_(row[c02.reserve - 1])
    );
  });

  // Tính lại giá vốn, lợi nhuận chịu thuế và Thuế TNDN theo từng dòng.
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const totalCost =
      FS02_POOL_num_(r[c02.cpxd - 1]) +
      FS02_POOL_num_(r[c02.selling - 1]) +
      FS02_POOL_num_(r[c02.operating - 1]) +
      FS02_POOL_num_(r[c02.gpmb - 1]) +
      FS02_POOL_num_(r[c02.htkt - 1]) +
      FS02_POOL_num_(r[c02.land - 1]) +
      FS02_POOL_num_(r[c02.reserve - 1]) +
      FS02_POOL_num_(r[c02.interest - 1]);
    const taxableProfit = Math.max(0, FS02_POOL_num_(r[c02.revenue - 1]) - totalCost);
    const cit = taxableProfit * FS02_POOL_rate_(r[c02.citRate - 1]);
    r[c02.totalCost - 1] = totalCost;
    r[c02.taxableProfit - 1] = taxableProfit;
    r[c02.cit - 1] = cit;
  }

  sh02.getRange(m02.headerRow + 1, 1, rows.length, width).setValues(rows);
  SpreadsheetApp.flush();

  const actualLand = FS02_POOL_sumColumn_(sh02, m02.headerRow + 1, c02.land);
  const actualReserve = FS02_POOL_sumColumn_(sh02, m02.headerRow + 1, c02.reserve);
  const actualInterest = FS02_POOL_sumColumn_(sh02, m02.headerRow + 1, c02.interest);
  const tol = 1;
  if (Math.abs(actualLand - sourceLand) > tol ||
      Math.abs(actualReserve - sourceReserve) > tol ||
      Math.abs(actualInterest - sourceInterest) > tol) {
    throw new Error(
      'Đối chiếu Sheet 02 chưa khớp nguồn: ' +
      'Đất lệch ' + (actualLand - sourceLand) + '; ' +
      'Dự phòng lệch ' + (actualReserve - sourceReserve) + '; ' +
      'Lãi vay lệch ' + (actualInterest - sourceInterest) + '.'
    );
  }
}

function FS02_POOL_reconcileColumn_(rows, colIndex, sourceTotal, weightFn) {
  const target = FS02_POOL_num_(sourceTotal);
  if (Math.abs(target) <= 0.5) {
    rows.forEach(r => r[colIndex] = 0);
    return;
  }

  let weights = rows.map(r => Math.max(0, FS02_POOL_num_(weightFn(r))));
  let totalWeight = weights.reduce((s, v) => s + v, 0);

  if (totalWeight <= 0) {
    // Chỉ phân bổ vào dòng có doanh thu hoặc có giá vốn đầu tư phát sinh.
    weights = rows.map(r => {
      const anyNumeric = r.some(v => typeof v === 'number' && v !== 0);
      return anyNumeric ? 1 : 0;
    });
    totalWeight = weights.reduce((s, v) => s + v, 0);
  }
  if (totalWeight <= 0) throw new Error('Không có trọng số để phân bổ nguồn ' + target + '.');

  let allocated = 0;
  let lastPositive = -1;
  for (let i = 0; i < rows.length; i++) {
    if (weights[i] > 0) lastPositive = i;
    const value = target * weights[i] / totalWeight;
    rows[i][colIndex] = value;
    allocated += value;
  }
  if (lastPositive >= 0) rows[lastPositive][colIndex] += target - allocated;
}

function FS02_POOL_detectHeader_(sheet, requiredNames) {
  const maxRows = Math.min(6, sheet.getLastRow());
  let best = null;
  for (let row = 1; row <= maxRows; row++) {
    const raw = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    const headers = raw.map(FS02_POOL_key_);
    const score = requiredNames.reduce((s, name) => s + (headers.indexOf(FS02_POOL_key_(name)) >= 0 ? 1 : 0), 0);
    if (!best || score > best.score) best = { headerRow: row, dataStartRow: row + 1, headers, score };
  }
  if (!best || best.score === 0) throw new Error('Không nhận diện được dòng tiêu đề tại ' + sheet.getName());
  return best;
}

function FS02_POOL_requireCol_(headers, names) {
  for (const name of names) {
    const i = headers.indexOf(FS02_POOL_key_(name));
    if (i >= 0) return i + 1;
  }
  throw new Error('Không tìm thấy cột: ' + names.join(' / '));
}

function FS02_POOL_sumNamed_(sheet, meta, name) {
  const col = FS02_POOL_requireCol_(meta.headers, [name]);
  return FS02_POOL_sumColumn_(sheet, meta.dataStartRow, col);
}

function FS02_POOL_sumColumn_(sheet, startRow, col) {
  if (sheet.getLastRow() < startRow) return 0;
  return sheet.getRange(startRow, col, sheet.getLastRow() - startRow + 1, 1)
    .getValues().reduce((s, r) => s + FS02_POOL_num_(r[0]), 0);
}

function FS02_POOL_rate_(value) {
  let n = FS02_POOL_num_(value);
  if (Math.abs(n) > 1) n /= 100;
  return Math.max(0, n);
}

function FS02_POOL_num_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  const s = String(value == null ? '' : value).trim().replace(/\s/g, '').replace(/\./g, '').replace(',', '.');
  const n = Number(s);
  return isFinite(n) ? n : 0;
}

function FS02_POOL_key_(value) {
  return String(value || '').toLowerCase().normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd')
    .replace(/[^a-z0-9_%/]+/g, ' ').replace(/\s+/g, ' ').trim();
}
