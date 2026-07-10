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

  try {
    if (!tech) throw new Error('Không tìm thấy sheet "01. Kỹ thuật".');
    let sh = ss.getSheetByName('05. Độ nhạy');
    if (!sh) throw new Error('Chưa có sheet "05. Độ nhạy" để lấy biên độ độ nhạy dự án.');

    const levels = FS05D_readProjectSensitivityLevels_(sh);
    if (!levels.length) {
      throw new Error('Không đọc được các mức biến động từ bảng độ nhạy NPV/IRR dự án hiện có.');
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
    const originalProducts = productRange.getValues();
    const originalCosts = costRange.getValues();
    const interestResults = [];
    const investmentResults = [];

    levels.forEach(change => {
      rateCell.setValue(originalRate * (1 + change));
      FS05C_runScenarioModel_();
      const m = FS05C_readEquityMetrics_();
      interestResults.push([change, originalRate * (1 + change), m.npv, m.irr]);
    });

    FS05D_restoreCell_(rateCell, rateFormula, rateValue);

    levels.forEach(change => {
      FS05C_applyInvestmentFactor_(
        tech, productBlock, costBlock,
        originalProducts, originalCosts, 1 + change
      );
      FS05C_runScenarioModel_();
      const m = FS05C_readEquityMetrics_();
      investmentResults.push([change, 1 + change, m.npv, m.irr]);
    });

    FS05D_restoreCell_(rateCell, rateFormula, rateValue);
    FS05D_restore_(productRange, productSnapshot);
    FS05D_restore_(costRange, costSnapshot);
    FS05C_runScenarioModel_();

    FS05C_writeResults_(sh, interestResults, investmentResults);
    sh.getRange('H8:K8').merge().setValue(
      'Biên độ lấy trực tiếp từ bảng độ nhạy NPV/IRR dự án hiện có'
    );
    ss.toast('Đã cập nhật độ nhạy vốn CSH theo cùng biên độ độ nhạy dự án.', 'FS - Độ nhạy', 8);
  } finally {
    try {
      if (rateCell) FS05D_restoreCell_(rateCell, rateFormula, rateValue);
      if (productRange && productSnapshot.length) FS05D_restore_(productRange, productSnapshot);
      if (costRange && costSnapshot.length) FS05D_restore_(costRange, costSnapshot);
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
  }
}

function FS05D_readProjectSensitivityLevels_(sheet) {
  const values = sheet.getDataRange().getValues();
  const candidates = [];

  values.forEach(row => row.forEach(v => {
    if (typeof v !== 'number' || !isFinite(v)) return;
    // Chỉ nhận các tỷ lệ biến động hợp lý, loại hệ số tiền và chỉ tiêu NPV/IRR.
    const rate = Math.abs(v) > 1 ? v / 100 : v;
    if (rate >= -0.50 && rate <= 0.50) candidates.push(rate);
  }));

  const unique = [...new Set(candidates.map(v => Number(v.toFixed(8))))]
    .filter(v => v !== 1)
    .sort((a, b) => a - b);

  // Ưu tiên dãy có cả âm, 0 và dương — đúng cấu trúc bảng độ nhạy dự án.
  const hasNegative = unique.some(v => v < 0);
  const hasZero = unique.some(v => Math.abs(v) < 1e-10);
  const hasPositive = unique.some(v => v > 0);
  return hasNegative && hasZero && hasPositive ? unique : [];
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
