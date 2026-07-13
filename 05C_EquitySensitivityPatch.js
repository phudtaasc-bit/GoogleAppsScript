/*************************************************
 * 05C_EquitySensitivityHelpers.gs
 * Helper dùng chung cho độ nhạy NPV/IRR vốn CSH.
 * Hàm chạy chính nằm tại 05D_EquitySensitivityAligned.js.
 *************************************************/

function FS05C_runScenarioModel_() {
  FS97_assertLandCostConfig_();

  // Khởi tạo Sheet 03 chỉ với chi phí/VAT; Sheet 04 là nguồn tính tài trợ duy nhất.
  FS_lapSheet03_CostOnly();
  FS_lapSheet02();
  FS_lapSheet04();
  SpreadsheetApp.flush();

  const result = FS_hoiTuLaiVay_();
  if (!result.converged) {
    throw new Error(
      'Kịch bản độ nhạy chưa hội tụ tài trợ sau ' + result.iterations +
      ' vòng; sai lệch còn lại ' + Math.round(result.maxDelta).toLocaleString('vi-VN') + ' đồng.'
    );
  }

  FS94_assertCITConsistency_();
  FS_lapSheet04A();
  FS_lapSheet00();
  FS93_runRegressionSuite_({ showAlert: false, throwOnError: true });
  SpreadsheetApp.flush();
}

function FS05C_readEquityMetrics_() {
  const ss = SpreadsheetApp.getActive();
  const sh00 = ss.getSheetByName('00. Tổng hợp');
  if (!sh00) throw new Error('Không tìm thấy sheet "00. Tổng hợp" sau khi chạy kịch bản.');

  return {
    npv: FS05C_num_(sh00.getRange('D36').getValue()) * 1000000000,
    irr: FS05C_rate_(sh00.getRange('D37').getValue())
  };
}

function FS05C_getBlock_(sheet, blockName) {
  const blockRow = FS05C_findExactRow_(sheet, blockName);
  if (!blockRow) throw new Error('Không tìm thấy block ' + blockName + '.');

  const headerRow = blockRow + 1;
  const display = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  let lastCol = 1;
  display.forEach((v, i) => { if (String(v || '').trim()) lastCol = i + 1; });

  const rows = [];
  let blank = 0;
  for (let r = blockRow + 2; r <= sheet.getLastRow(); r++) {
    const first = String(sheet.getRange(r, 1).getDisplayValue() || '').trim();
    if (FS05C_isBlock_(first)) break;
    const row = sheet.getRange(r, 1, 1, lastCol).getValues()[0];
    if (!row.some(v => String(v === null ? '' : v).trim() !== '')) {
      blank++;
      if (blank >= 3) break;
      continue;
    }
    blank = 0;
    rows.push(row);
  }

  if (!rows.length) throw new Error('Block ' + blockName + ' không có dữ liệu.');
  return { blockRow, headerRow, startRow: blockRow + 2, lastCol, rows };
}

function FS05C_findExactRow_(sheet, text) {
  const values = sheet.getDataRange().getDisplayValues();
  const target = FS05C_key_(text);
  for (let r = 0; r < values.length; r++) {
    if (values[r].some(v => FS05C_key_(v) === target)) return r + 1;
  }
  return 0;
}

function FS05C_isBlock_(value) {
  return ['THONG_TIN_CHUNG', 'CHI_PHI_CHUNG', 'SAN_PHAM', 'KE_HOACH_BAN_THU_TIEN', 'TIEN_DO_CHI_PHI']
    .some(x => FS05C_key_(value) === FS05C_key_(x));
}

function FS05C_containsAny_(key, terms) {
  return terms.some(x => key.indexOf(FS05C_key_(x)) >= 0);
}

function FS05C_num_(value) {
  const n = Number(value);
  return isFinite(n) ? n : 0;
}

function FS05C_rate_(value) {
  const n = Number(value);
  if (!isFinite(n)) return 0;
  return n > 1 ? n / 100 : n;
}

function FS05C_key_(value) {
  return String(value || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}
