/*************************************************
 * 05D_EquitySensitivityAligned.gs
 * Độ nhạy NPV/IRR vốn CSH dùng cùng biên độ với độ nhạy dự án hiện có.
 *************************************************/

function FS05D_DoNhay_CSH_LaiSuat_VonDauTu() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName('01. Kỹ thuật');
  let rateCell, productRange, costRange;
  let rateFormula = '', rateValue = 0, productSnapshot = [], costSnapshot = [];
  let interestResults = [], investmentResults = [];
  let primaryError = null;

  try {
    if (!tech) throw new Error('Không tìm thấy sheet "01. Kỹ thuật".');
    const sh = ss.getSheetByName('05. Độ nhạy');
    if (!sh) throw new Error('Chưa có sheet "05. Độ nhạy" để lấy biên độ độ nhạy dự án.');

    const levels = FS05D_readProjectSensitivityLevels_(sh);
    if (!levels.length) {
      throw new Error(
        'Không đọc được dãy biến động có đủ mức âm, 0 và dương từ bảng độ nhạy NPV/IRR dự án hiện có.'
      );
    }

    const rateRow = FS05C_findExactRow_(tech, 'Lãi suất vay năm');
    if (!rateRow) throw new Error('Không tìm thấy "Lãi suất vay năm" tại 01. Kỹ thuật.');

    const productBlock = FS05C_getBlock_(tech, 'SAN_PHAM');
    const costBlock = FS05C_getBlock_(tech, 'CHI_PHI_CHUNG');
    rateCell = tech.getRange(rateRow, 2);
    productRange = tech.getRange(productBlock.startRow, 1, productBlock.rows.length, productBlock.lastCol);
    costRange = tech.getRange(costBlock.startRow, 1, costBlock.rows.length, costBlock.lastCol);

    rateFormula = rateCell.getFormula();
    rateValue = rateCell.getValue();
    productSnapshot = FS05D_snapshot_(productRange);
    costSnapshot = FS05D_snapshot_(costRange);

    const originalRate = FS05C_rate_(rateValue);

    ss.toast('Đang tính độ nhạy vốn CSH theo lãi suất...', 'FS - Độ nhạy', 5);
    levels.forEach(change => {
      rateCell.setValue(originalRate * (1 + change));
      FS05C_runScenarioModel_();
      const m = FS05C_readEquityMetrics_();
      interestResults.push([change, originalRate * (1 + change), m.npv, m.irr]);
    });

    FS05D_restoreCell_(rateCell, rateFormula, rateValue);

    ss.toast('Đang tính độ nhạy vốn CSH theo vốn đầu tư...', 'FS - Độ nhạy', 5);
    levels.forEach(change => {
      FS05D_applyInvestmentFactor_(productRange, costRange, productSnapshot, costSnapshot, 1 + change);
      FS05C_runScenarioModel_();
      const m = FS05C_readEquityMetrics_();
      investmentResults.push([change, 1 + change, m.npv, m.irr]);
    });
  } catch (err) {
    primaryError = err;
  } finally {
    try {
      if (rateCell) FS05D_restoreCell_(rateCell, rateFormula, rateValue);
      if (productRange && productSnapshot.length) FS05D_restore_(productRange, productSnapshot);
      if (costRange && costSnapshot.length) FS05D_restore_(costRange, costSnapshot);
      SpreadsheetApp.flush();

      // Luôn trả cả đầu vào và kết quả mô hình về trạng thái cơ sở.
      if (rateCell && productRange && costRange) FS05C_runScenarioModel_();
    } catch (restoreErr) {
      if (!primaryError) primaryError = restoreErr;
      else primaryError.message += '\nLỗi khi phục hồi mô hình cơ sở: ' + restoreErr.message;
    } finally {
      lock.releaseLock();
    }
  }

  if (primaryError) throw primaryError;

  const sh = ss.getSheetByName('05. Độ nhạy');
  FS05D_writeResults_(sh, interestResults, investmentResults);
  ss.toast('Đã cập nhật độ nhạy vốn CSH theo cùng biên độ độ nhạy dự án.', 'FS - Độ nhạy', 8);
}

function FS05D_readProjectSensitivityLevels_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values.length) return [];

  const candidates = [];
  values.forEach(row => FS05D_collectSequences_(row).forEach(x => candidates.push(x)));

  const width = Math.max.apply(null, values.map(r => r.length));
  for (let c = 0; c < width; c++) {
    const column = values.map(r => r[c]);
    FS05D_collectSequences_(column).forEach(x => candidates.push(x));
  }

  candidates.sort((a, b) => {
    if (b.length !== a.length) return b.length - a.length;
    return FS05D_stepVariance_(a) - FS05D_stepVariance_(b);
  });

  return candidates.length ? candidates[0] : [];
}

function FS05D_collectSequences_(items) {
  const out = [];
  let current = [];

  const flush = () => {
    if (current.length >= 3) {
      const normalized = FS05D_normalizeSequence_(current);
      if (normalized.length) out.push(normalized);
    }
    current = [];
  };

  items.forEach(v => {
    const rate = FS05D_asSensitivityRate_(v);
    if (rate === null) flush();
    else current.push(rate);
  });
  flush();
  return out;
}

function FS05D_asSensitivityRate_(value) {
  if (typeof value !== 'number' || !isFinite(value)) return null;
  let rate = value;
  if (Math.abs(rate) > 1 && Math.abs(rate) <= 50) rate /= 100;
  if (rate < -0.50 || rate > 0.50) return null;
  return Number(rate.toFixed(10));
}

function FS05D_normalizeSequence_(sequence) {
  const unique = sequence.filter((v, i) => i === 0 || Math.abs(v - sequence[i - 1]) > 1e-10);
  if (unique.length < 3 || unique.length > 15) return [];

  const increasing = unique.every((v, i) => i === 0 || v > unique[i - 1]);
  const decreasing = unique.every((v, i) => i === 0 || v < unique[i - 1]);
  if (!increasing && !decreasing) return [];

  const sorted = increasing ? unique.slice() : unique.slice().reverse();
  const hasNegative = sorted.some(v => v < -1e-10);
  const hasZero = sorted.some(v => Math.abs(v) <= 1e-10);
  const hasPositive = sorted.some(v => v > 1e-10);
  if (!hasNegative || !hasZero || !hasPositive) return [];

  return sorted;
}

function FS05D_stepVariance_(sequence) {
  if (sequence.length < 3) return Infinity;
  const steps = [];
  for (let i = 1; i < sequence.length; i++) steps.push(sequence[i] - sequence[i - 1]);
  const mean = steps.reduce((s, v) => s + v, 0) / steps.length;
  return steps.reduce((s, v) => s + Math.pow(v - mean, 2), 0) / steps.length;
}

function FS05D_snapshot_(range) {
  const values = range.getValues();
  const formulas = range.getFormulas();
  return values.map((row, r) => row.map((v, c) => formulas[r][c] || v));
}

function FS05D_restore_(range, matrix) {
  if (matrix && matrix.length) range.setValues(matrix);
}

function FS05D_restoreCell_(cell, formula, value) {
  if (formula) cell.setFormula(formula);
  else cell.setValue(value);
}

function FS05D_applyInvestmentFactor_(productRange, costRange, productSnapshot, costSnapshot, factor) {
  const products = productSnapshot.map(r => r.slice());
  products.forEach(r => {
    if (typeof r[5] === 'number' && isFinite(r[5])) r[5] *= factor;
  });

  const costs = costSnapshot.map(r => r.slice());
  costs.forEach(r => {
    const name = FS05C_key_(r[0]);
    const excluded = FS05C_containsAny_(name, [
      'chi phi ban hang',
      'chi phi van hanh',
      'chi phi du phong',
      'chi phi lai vay'
    ]);

    if (!excluded && typeof r[1] === 'number' && isFinite(r[1])) {
      r[1] *= factor;
      if (typeof r[3] === 'number' && isFinite(r[3])) r[3] *= factor;
    }
  });

  productRange.setValues(products);
  costRange.setValues(costs);
}

function FS05D_writeResults_(sheet, interestResults, investmentResults) {
  const n1 = interestResults.length;
  const n2 = investmentResults.length;
  const secondTitleRow = 4 + n1;
  const secondHeaderRow = secondTitleRow + 1;
  const secondDataRow = secondHeaderRow + 1;
  const lastRow = secondDataRow + n2 - 1;

  if (sheet.getMaxColumns() < 12) sheet.insertColumnsAfter(sheet.getMaxColumns(), 12 - sheet.getMaxColumns());
  if (sheet.getMaxRows() < lastRow) sheet.insertRowsAfter(sheet.getMaxRows(), lastRow - sheet.getMaxRows());

  sheet.getRange(1, 8, Math.max(lastRow, 1), 4).clearContent().clearFormat().breakApart();

  sheet.getRange(1, 8, 1, 4).merge().setValue('ĐỘ NHẠY NPV/IRR VỐN CSH THEO LÃI SUẤT VAY');
  sheet.getRange(2, 8, 1, 4).setValues([['Biến động', 'Lãi suất vay năm', 'NPV vốn CSH', 'IRR vốn CSH']]);
  if (n1) sheet.getRange(3, 8, n1, 4).setValues(interestResults);

  sheet.getRange(secondTitleRow, 8, 1, 4).merge().setValue('ĐỘ NHẠY NPV/IRR VỐN CSH THEO VỐN ĐẦU TƯ');
  sheet.getRange(secondHeaderRow, 8, 1, 4).setValues([['Biến động', 'Hệ số vốn đầu tư', 'NPV vốn CSH', 'IRR vốn CSH']]);
  if (n2) sheet.getRange(secondDataRow, 8, n2, 4).setValues(investmentResults);

  sheet.getRange(lastRow + 2, 8, 1, 4).merge().setValue(
    'Biên độ lấy trực tiếp từ dãy biến động của bảng độ nhạy NPV/IRR dự án hiện có'
  );

  [sheet.getRange(1, 8, 1, 4), sheet.getRange(secondTitleRow, 8, 1, 4)]
    .forEach(r => r.setFontWeight('bold').setHorizontalAlignment('center'));
  [sheet.getRange(2, 8, 1, 4), sheet.getRange(secondHeaderRow, 8, 1, 4)]
    .forEach(r => r.setFontWeight('bold').setHorizontalAlignment('center').setWrap(true));

  if (n1) {
    sheet.getRange(3, 8, n1, 2).setNumberFormat('0.00%');
    sheet.getRange(3, 10, n1, 1).setNumberFormat('#,##0');
    sheet.getRange(3, 11, n1, 1).setNumberFormat('0.00%');
  }
  if (n2) {
    sheet.getRange(secondDataRow, 8, n2, 1).setNumberFormat('0.00%');
    sheet.getRange(secondDataRow, 9, n2, 1).setNumberFormat('0.00x');
    sheet.getRange(secondDataRow, 10, n2, 1).setNumberFormat('#,##0');
    sheet.getRange(secondDataRow, 11, n2, 1).setNumberFormat('0.00%');
  }

  sheet.autoResizeColumns(8, 4);
}
