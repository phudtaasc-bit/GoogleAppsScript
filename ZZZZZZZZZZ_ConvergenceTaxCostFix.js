/*************************************************
 * ZZZZZZZZZZ_ConvergenceTaxCostFix.js
 *
 * Mục tiêu:
 * 1) Giữ lãi vay trong giá vốn tính thuế.
 * 2) Chi phí vận hành và chi phí bảo trì đều thuộc giá vốn tính thuế.
 * 3) Chạy lặp Sheet 02 <-> Sheet 04 đến khi tổng lãi vay và Thuế TNDN hội tụ.
 * 4) Không ép cân đối Sheet 00; nếu chưa hội tụ hoặc audit còn FAIL thì báo lỗi.
 *************************************************/

const FS_CONV_BASE_LAP_SHEET02_ = FS_lapSheet02;

FS_lapSheet02 = function() {
  FS_CONV_BASE_LAP_SHEET02_();
  FS_CONV_rebuildTaxCostIncludingOpMaint_();
};

/**
 * Tính lại giá vốn tính thuế theo đúng các cấu phần hiện có trên Sheet 02.
 * Vận hành và bảo trì chỉ được cộng đúng một lần.
 */
function FS_CONV_rebuildTaxCostIncludingOpMaint_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('02. Doanh thu');
  if (!sh) throw new Error('Không tìm thấy Sheet 02. Doanh thu.');

  const headerRow = FS_CONV_detectHeaderRow_(sh, [
    'Tổng giá vốn tính thuế',
    'Lợi nhuận chịu thuế',
    'Thuế TNDN tạm tính'
  ]);
  const headers = sh.getRange(headerRow, 1, 1, sh.getLastColumn()).getDisplayValues()[0];

  const col = names => FS_CONV_findCol_(headers, names);
  const required = names => {
    const c = col(names);
    if (c < 1) throw new Error('Sheet 02 thiếu cột: ' + names.join(' / '));
    return c;
  };

  const monthCol = required(['Tháng số']);
  const revenueCol = required(['Tổng doanh thu trước VAT']);
  const taxRateCol = required(['Thuế TNDN %']);
  const totalTaxCostCol = required(['Tổng giá vốn tính thuế']);
  const taxableProfitCol = required(['Lợi nhuận chịu thuế']);
  const citCol = required(['Thuế TNDN tạm tính']);

  const componentCols = [
    col(['CP XD/TB trực tiếp trước VAT']),
    col(['Chi phí bán hàng trước VAT']),
    col(['Chi phí vận hành thuê trước VAT', 'Chi phí vận hành trước VAT']),
    col(['Chi phí bảo trì trước VAT']),
    col(['Chi phí GPMB phân bổ trước VAT']),
    col(['Chi phí HTKT phân bổ trước VAT']),
    col(['Tiền SDĐ/thuê đất phân bổ trước VAT']),
    col(['Chi phí dự phòng phân bổ trước VAT']),
    col(['Chi phí lãi vay phân bổ'])
  ].filter(c => c > 0);

  if (!componentCols.length) throw new Error('Không tìm thấy các cột cấu phần giá vốn tại Sheet 02.');

  const startRow = headerRow + 1;
  const lastRow = sh.getLastRow();
  if (lastRow < startRow) return;
  const n = lastRow - startRow + 1;
  const width = sh.getLastColumn();
  const data = sh.getRange(startRow, 1, n, width).getValues();

  const outCost = [];
  const outProfit = [];
  const outCit = [];

  data.forEach(r => {
    const month = Number(r[monthCol - 1]) || 0;
    if (!month) {
      outCost.push(['']);
      outProfit.push(['']);
      outCit.push(['']);
      return;
    }

    const totalTaxCost = componentCols.reduce((s, c) => s + (Number(r[c - 1]) || 0), 0);
    const revenue = Number(r[revenueCol - 1]) || 0;
    const taxRate = FS_CONV_rate_(r[taxRateCol - 1]);
    const taxableProfit = Math.max(0, revenue - totalTaxCost);
    const cit = taxableProfit * taxRate;

    outCost.push([totalTaxCost]);
    outProfit.push([taxableProfit]);
    outCit.push([cit]);
  });

  sh.getRange(startRow, totalTaxCostCol, n, 1).setValues(outCost);
  sh.getRange(startRow, taxableProfitCol, n, 1).setValues(outProfit);
  sh.getRange(startRow, citCol, n, 1).setValues(outCit);
  SpreadsheetApp.flush();
}

/**
 * Chạy toàn bộ mô hình có cập nhật lãi vay theo cơ chế lặp hội tụ.
 * Điều kiện hội tụ đồng thời:
 * - Tổng lãi vay Sheet 04 thay đổi không quá 1.000 đồng.
 * - Tổng Thuế TNDN Sheet 02 thay đổi không quá 1.000 đồng.
 */
function FS_chayToanBoMoHinh_CoLaiVay() {
  const MAX_ITER = 10;
  const TOL = 1000;

  if (typeof FS_taoKyThuatTuDauVao !== 'function') throw new Error('Thiếu hàm FS_taoKyThuatTuDauVao.');
  if (typeof FS_lapSheet03 !== 'function') throw new Error('Thiếu hàm FS_lapSheet03.');
  if (typeof FS_lapSheet02 !== 'function') throw new Error('Thiếu hàm FS_lapSheet02.');
  if (typeof FS_lapSheet04 !== 'function') throw new Error('Thiếu hàm FS_lapSheet04.');

  FS_taoKyThuatTuDauVao();
  FS_lapSheet03();

  let prevInterest = null;
  let prevCit = null;
  let converged = false;
  let finalInterest = 0;
  let finalCit = 0;
  let iterDone = 0;

  for (let iter = 1; iter <= MAX_ITER; iter++) {
    FS_lapSheet02();
    SpreadsheetApp.flush();

    FS_lapSheet04();
    SpreadsheetApp.flush();

    // Giữ tương thích với mô hình cũ nếu hàm đồng bộ 04 -> 03 còn tồn tại.
    if (typeof FS03_capNhatNguonVonTuSheet04 === 'function') {
      FS03_capNhatNguonVonTuSheet04();
      SpreadsheetApp.flush();
    }

    finalInterest = FS_CONV_sumNamed_('04. Dòng tiền & Lợi nhuận', [
      'Lãi vay vốn hóa', 'Lãi vay'
    ]);
    finalCit = FS_CONV_sumNamed_('02. Doanh thu', ['Thuế TNDN tạm tính']);
    iterDone = iter;

    if (prevInterest !== null && prevCit !== null &&
        Math.abs(finalInterest - prevInterest) <= TOL &&
        Math.abs(finalCit - prevCit) <= TOL) {
      converged = true;
      break;
    }

    prevInterest = finalInterest;
    prevCit = finalCit;
  }

  if (!converged) {
    throw new Error(
      'Mô hình chưa hội tụ sau ' + MAX_ITER + ' vòng. ' +
      'Tổng lãi vay cuối: ' + finalInterest + '; Thuế TNDN cuối: ' + finalCit + '. ' +
      'Không tiếp tục lập Sheet 00.'
    );
  }

  // Chạy lại 02 và 04 lần cuối trên bộ lãi vay đã hội tụ.
  FS_lapSheet02();
  SpreadsheetApp.flush();
  FS_lapSheet04();
  SpreadsheetApp.flush();

  if (typeof FS03_capNhatNguonVonTuSheet04 === 'function') {
    FS03_capNhatNguonVonTuSheet04();
    SpreadsheetApp.flush();
  }

  if (typeof FS_lapSheet04A === 'function') FS_lapSheet04A();
  if (typeof FS_lapSheet00 === 'function') FS_lapSheet00();
  if (typeof FS99Q_chayAuditReconciliation === 'function') FS99Q_chayAuditReconciliation();

  SpreadsheetApp.getUi().alert(
    'Đã chạy toàn bộ mô hình và hội tụ sau ' + iterDone + ' vòng.\n' +
    'Tổng lãi vay: ' + FS_CONV_format_(finalInterest) + ' đồng.\n' +
    'Tổng Thuế TNDN: ' + FS_CONV_format_(finalCit) + ' đồng.'
  );
}

function FS_CONV_sumNamed_(sheetName, names) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh) throw new Error('Không tìm thấy sheet ' + sheetName + '.');
  const headerRow = FS_CONV_detectHeaderRow_(sh, names);
  const headers = sh.getRange(headerRow, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  const c = FS_CONV_findCol_(headers, names);
  if (c < 1) throw new Error('Không tìm thấy cột ' + names.join(' / ') + ' tại ' + sheetName + '.');
  const start = headerRow + 1;
  if (sh.getLastRow() < start) return 0;
  return sh.getRange(start, c, sh.getLastRow() - start + 1, 1)
    .getValues().reduce((s, r) => s + (Number(r[0]) || 0), 0);
}

function FS_CONV_detectHeaderRow_(sh, requiredNames) {
  const maxRows = Math.min(6, sh.getLastRow());
  let best = null;
  for (let r = 1; r <= maxRows; r++) {
    const headers = sh.getRange(r, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
    const hits = requiredNames.reduce((s, n) => s + (FS_CONV_findCol_(headers, [n]) > 0 ? 1 : 0), 0);
    if (!best || hits > best.hits) best = { row: r, hits };
  }
  if (!best || best.hits < 1) throw new Error('Không nhận diện được dòng tiêu đề tại ' + sh.getName() + '.');
  return best.row;
}

function FS_CONV_findCol_(headers, names) {
  const normalized = headers.map(FS_CONV_key_);
  for (const name of names) {
    const idx = normalized.indexOf(FS_CONV_key_(name));
    if (idx >= 0) return idx + 1;
  }
  return -1;
}

function FS_CONV_rate_(v) {
  let n = Number(v);
  if (!isFinite(n)) n = 0;
  if (Math.abs(n) > 1) n /= 100;
  return Math.max(0, n);
}

function FS_CONV_key_(v) {
  return String(v || '').toLowerCase().normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function FS_CONV_format_(v) {
  return Math.round(Number(v) || 0).toLocaleString('vi-VN');
}
