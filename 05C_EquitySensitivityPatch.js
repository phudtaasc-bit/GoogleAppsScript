/*************************************************
 * 05C_EquitySensitivityPatch.gs
 * Bổ sung độ nhạy NPV/IRR vốn CSH theo:
 * 1) Tăng/giảm lãi suất vay.
 * 2) Tăng/giảm vốn đầu tư.
 *
 * Không ghi đè các bảng độ nhạy hiện có; kết quả đặt từ cột H.
 * Mọi dữ liệu đầu vào được snapshot và khôi phục sau khi chạy.
 *************************************************/

const FS05C_LEVELS = [-0.10, -0.05, 0, 0.05, 0.10];

function FS05C_DoNhay_CSH_LaiSuat_VonDauTu() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName('01. Kỹ thuật');
  if (!tech) throw new Error('Không tìm thấy sheet "01. Kỹ thuật".');

  let sh = ss.getSheetByName('05. Độ nhạy');
  if (!sh) sh = ss.insertSheet('05. Độ nhạy');

  const rateRow = FS05C_findExactRow_(tech, 'Lãi suất vay năm');
  if (!rateRow) throw new Error('Không tìm thấy "Lãi suất vay năm" tại 01. Kỹ thuật.');

  const productBlock = FS05C_getBlock_(tech, 'SAN_PHAM');
  const costBlock = FS05C_getBlock_(tech, 'CHI_PHI_CHUNG');
  const rateCell = tech.getRange(rateRow, 2);

  const originalRate = FS05C_rate_(rateCell.getValue());
  const originalProductRows = productBlock.rows.map(r => r.slice());
  const originalCostRows = costBlock.rows.map(r => r.slice());

  const interestResults = [];
  const investmentResults = [];

  try {
    ss.toast('Đang tính độ nhạy vốn CSH theo lãi suất...', 'FS - Độ nhạy', 5);

    FS05C_LEVELS.forEach(change => {
      rateCell.setValue(originalRate * (1 + change));
      FS05C_runScenarioModel_();
      const metrics = FS05C_readEquityMetrics_();
      interestResults.push([change, originalRate * (1 + change), metrics.npv, metrics.irr]);
    });

    // Khôi phục lãi suất trước khi chạy nhóm vốn đầu tư.
    rateCell.setValue(originalRate);

    ss.toast('Đang tính độ nhạy vốn CSH theo vốn đầu tư...', 'FS - Độ nhạy', 5);

    FS05C_LEVELS.forEach(change => {
      FS05C_applyInvestmentFactor_(tech, productBlock, costBlock, originalProductRows, originalCostRows, 1 + change);
      FS05C_runScenarioModel_();
      const metrics = FS05C_readEquityMetrics_();
      investmentResults.push([change, 1 + change, metrics.npv, metrics.irr]);
    });
  } finally {
    // Khôi phục tuyệt đối dữ liệu nguồn, kể cả khi một kịch bản phát sinh lỗi.
    rateCell.setValue(originalRate);
    if (originalProductRows.length) {
      tech.getRange(productBlock.startRow, 1, originalProductRows.length, productBlock.lastCol)
        .setValues(originalProductRows);
    }
    if (originalCostRows.length) {
      tech.getRange(costBlock.startRow, 1, originalCostRows.length, costBlock.lastCol)
        .setValues(originalCostRows);
    }

    // Trả mô hình về trạng thái cơ sở.
    FS05C_runScenarioModel_();
    lock.releaseLock();
  }

  FS05C_writeResults_(sh, interestResults, investmentResults);
  ss.toast('Đã cập nhật độ nhạy NPV/IRR vốn CSH.', 'FS - Độ nhạy', 8);
}

function FS05C_runScenarioModel_() {
  FS97_assertLandCostConfig_();
  FS_lapSheet03_Patched();
  FS_lapSheet02();
  FS_lapSheet04();
  SpreadsheetApp.flush();

  const result = FS_hoiTuLaiVay_();
  if (!result.converged) {
    throw new Error(
      'Kịch bản độ nhạy chưa hội tụ lãi vay sau ' + result.iterations +
      ' vòng; sai lệch còn lại ' + Math.round(result.maxDelta).toLocaleString('vi-VN') + ' đồng.'
    );
  }

  FS94_assertCITConsistency_();
  FS_lapSheet04A();
  if (typeof FS_lapSheet00_Patched === 'function') {
    FS_lapSheet00_Patched();
  } else {
    FS_lapSheet00();
  }
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

function FS05C_applyInvestmentFactor_(tech, productBlock, costBlock, originalProducts, originalCosts, factor) {
  const products = originalProducts.map(r => r.slice());
  products.forEach(r => {
    // SAN_PHAM cột F: CPXD/m².
    if (typeof r[5] === 'number' && isFinite(r[5])) r[5] = r[5] * factor;
  });

  const costs = originalCosts.map(r => r.slice());
  costs.forEach(r => {
    const name = FS05C_key_(r[0]);
    const excluded = FS05C_containsAny_(name, [
      'chi phi ban hang',
      'chi phi van hanh',
      'chi phi du phong',
      'chi phi lai vay'
    ]);

    // CHI_PHI_CHUNG cột B: giá trị trước VAT tuyệt đối.
    // Không nhân các dòng tỷ lệ hoặc chi phí phụ thuộc doanh thu.
    if (!excluded && typeof r[1] === 'number' && isFinite(r[1])) {
      r[1] = r[1] * factor;
      // Nếu cột D là giá trị sau VAT tuyệt đối thì nhân đồng bộ.
      if (typeof r[3] === 'number' && isFinite(r[3])) r[3] = r[3] * factor;
    }
  });

  if (products.length) {
    tech.getRange(productBlock.startRow, 1, products.length, productBlock.lastCol).setValues(products);
  }
  if (costs.length) {
    tech.getRange(costBlock.startRow, 1, costs.length, costBlock.lastCol).setValues(costs);
  }
}

function FS05C_writeResults_(sh, interestResults, investmentResults) {
  if (sh.getMaxColumns() < 12) {
    sh.insertColumnsAfter(sh.getMaxColumns(), 12 - sh.getMaxColumns());
  }
  if (sh.getMaxRows() < 20) {
    sh.insertRowsAfter(sh.getMaxRows(), 20 - sh.getMaxRows());
  }

  sh.getRange('H1:K20').clearContent().clearFormat();

  sh.getRange('H1:K1').merge().setValue('ĐỘ NHẠY NPV/IRR VỐN CSH THEO LÃI SUẤT VAY');
  sh.getRange('H2:K2').setValues([[
    'Biến động', 'Lãi suất vay năm', 'NPV vốn CSH', 'IRR vốn CSH'
  ]]);
  if (interestResults.length) sh.getRange(3, 8, interestResults.length, 4).setValues(interestResults);

  sh.getRange('H10:K10').merge().setValue('ĐỘ NHẠY NPV/IRR VỐN CSH THEO VỐN ĐẦU TƯ');
  sh.getRange('H11:K11').setValues([[
    'Biến động', 'Hệ số vốn đầu tư', 'NPV vốn CSH', 'IRR vốn CSH'
  ]]);
  if (investmentResults.length) sh.getRange(12, 8, investmentResults.length, 4).setValues(investmentResults);

  ['H1:K1', 'H10:K10'].forEach(a1 => {
    sh.getRange(a1).setFontWeight('bold').setHorizontalAlignment('center');
  });
  ['H2:K2', 'H11:K11'].forEach(a1 => {
    sh.getRange(a1).setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
  });

  sh.getRange('H3:I7').setNumberFormat('0.00%');
  sh.getRange('J3:J7').setNumberFormat('#,##0');
  sh.getRange('K3:K7').setNumberFormat('0.00%');
  sh.getRange('H12:H16').setNumberFormat('0.00%');
  sh.getRange('I12:I16').setNumberFormat('0.00x');
  sh.getRange('J12:J16').setNumberFormat('#,##0');
  sh.getRange('K12:K16').setNumberFormat('0.00%');
  sh.autoResizeColumns(8, 4);
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
