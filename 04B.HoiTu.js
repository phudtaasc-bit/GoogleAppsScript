const FSCONV_CFG = Object.freeze({
  TECH: '01A. Kỹ thuật',
  MAX_ITERATIONS: 150,
  TOLERANCE: 1,
  TOTAL_INTEREST_TOLERANCE: 1,
  DAMPING: 0.5
});

function FS_hoiTuTaiTro() {
  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName(FSCONV_CFG.TECH);
  if (!tech) throw new Error('Thiếu sheet "01A. Kỹ thuật".');

  const months = Math.max(0, FSCONV_num_(FSCONV_readInfo_(tech, 'Số tháng mô hình')));
  if (!months) throw new Error('Số tháng mô hình phải lớn hơn 0.');

  let assumedInterest = Array(months).fill(0);
  let accountingRows = [];
  let financeResult = null;
  let converged = false;
  let maxDiff = Infinity;
  let totalInterestDiff = Infinity;
  let iteration = 0;

  for (iteration = 1; iteration <= FSCONV_CFG.MAX_ITERATIONS; iteration++) {
    accountingRows = FS03A_build_(assumedInterest, false);
    const taxByMonth = FSCONV_taxByMonth_(accountingRows, months);
    financeResult = FS04_build_(taxByMonth, false);

    const calculatedInterest = financeResult.interestByMonth;
    maxDiff = FSCONV_maxDiff_(assumedInterest, calculatedInterest);
    totalInterestDiff = FSCONV_totalDiff_(assumedInterest, calculatedInterest);

    if (
      maxDiff <= FSCONV_CFG.TOLERANCE &&
      totalInterestDiff <= FSCONV_CFG.TOTAL_INTEREST_TOLERANCE
    ) {
      assumedInterest = calculatedInterest.slice();
      converged = true;
      break;
    }

    assumedInterest = assumedInterest.map((oldValue, index) =>
      oldValue * (1 - FSCONV_CFG.DAMPING) +
      calculatedInterest[index] * FSCONV_CFG.DAMPING
    );
  }

  if (!converged) {
    FSCONV_writeStatus_(
      false,
      iteration - 1,
      maxDiff,
      totalInterestDiff,
      assumedInterest,
      financeResult
    );

    throw new Error(
      'Mô hình 03A–04 không hội tụ sau ' + FSCONV_CFG.MAX_ITERATIONS +
      ' vòng. Sai số lãi vay lớn nhất: ' +
      Math.round(maxDiff).toLocaleString('vi-VN') +
      ' đồng; chênh lệch tổng lãi vay: ' +
      Math.round(totalInterestDiff).toLocaleString('vi-VN') + ' đồng.'
    );
  }

  // Build lần cuối bằng đúng mảng lãi vay đã hội tụ.
  accountingRows = FS03A_build_(assumedInterest, false);
  const finalTaxByMonth = FSCONV_taxByMonth_(accountingRows, months);
  financeResult = FS04_build_(finalTaxByMonth, false);

  const finalMaxDiff = FSCONV_maxDiff_(assumedInterest, financeResult.interestByMonth);
  const finalTotalDiff = FSCONV_totalDiff_(assumedInterest, financeResult.interestByMonth);
  const accountingInterest = FSCONV_accountingInterest_(accountingRows);
  const financeInterest = financeResult.interestByMonth.reduce(
    (sum, value) => sum + FSCONV_num_(value),
    0
  );
  const sheetInterestDiff = Math.abs(accountingInterest - financeInterest);

  if (
    finalMaxDiff > FSCONV_CFG.TOLERANCE ||
    finalTotalDiff > FSCONV_CFG.TOTAL_INTEREST_TOLERANCE ||
    sheetInterestDiff > FSCONV_CFG.TOTAL_INTEREST_TOLERANCE
  ) {
    FSCONV_writeStatus_(
      false,
      iteration,
      finalMaxDiff,
      Math.max(finalTotalDiff, sheetInterestDiff),
      assumedInterest,
      financeResult
    );

    throw new Error(
      'Kết quả cuối chưa đối chiếu được lãi vay giữa 03A và 04. ' +
      'Tổng lãi vay 03A: ' + Math.round(accountingInterest).toLocaleString('vi-VN') +
      '; tổng lãi vay 04: ' + Math.round(financeInterest).toLocaleString('vi-VN') +
      '; chênh lệch: ' + Math.round(sheetInterestDiff).toLocaleString('vi-VN') + ' đồng.'
    );
  }

  FS03A_write_(ss, accountingRows);
  FS04_write_(ss, financeResult.rows);
  FSCONV_writeStatus_(
    true,
    iteration,
    finalMaxDiff,
    sheetInterestDiff,
    assumedInterest,
    financeResult
  );

  return {
    converged: true,
    iterations: iteration,
    maxDiff: finalMaxDiff,
    totalInterestDiff: sheetInterestDiff,
    accountingRows,
    financeRows: financeResult.rows,
    interestByMonth: financeResult.interestByMonth
  };
}

function FSCONV_taxByMonth_(accountingRows, months) {
  const result = Array(months).fill(0);
  accountingRows.forEach(row => {
    const monthNo = FSCONV_num_(row[0]);
    if (monthNo >= 1 && monthNo <= months) {
      result[monthNo - 1] += FSCONV_num_(row[21]);
    }
  });
  return result;
}

function FSCONV_accountingInterest_(accountingRows) {
  return accountingRows.reduce(
    (sum, row) => sum + FSCONV_num_(row[11]),
    0
  );
}

function FSCONV_maxDiff_(left, right) {
  let max = 0;
  const length = Math.max(left.length, right.length);
  for (let index = 0; index < length; index++) {
    max = Math.max(
      max,
      Math.abs(FSCONV_num_(left[index]) - FSCONV_num_(right[index]))
    );
  }
  return max;
}

function FSCONV_totalDiff_(left, right) {
  const leftTotal = left.reduce((sum, value) => sum + FSCONV_num_(value), 0);
  const rightTotal = right.reduce((sum, value) => sum + FSCONV_num_(value), 0);
  return Math.abs(leftTotal - rightTotal);
}

function FSCONV_writeStatus_(
  converged,
  iterations,
  maxDiff,
  totalInterestDiff,
  interest,
  financeResult
) {
  const rows = financeResult && financeResult.rows ? financeResult.rows : [];
  const last = rows.length ? rows[rows.length - 1] : [];
  const totalInterest = interest.reduce((sum, value) => sum + FSCONV_num_(value), 0);

  PropertiesService.getDocumentProperties().setProperty('FS_CONVERGENCE', JSON.stringify({
    converged,
    iterations,
    maxDiff,
    totalInterestDiff,
    tolerance: FSCONV_CFG.TOLERANCE,
    totalInterestTolerance: FSCONV_CFG.TOTAL_INTEREST_TOLERANCE,
    damping: FSCONV_CFG.DAMPING,
    totalInterest,
    closingDebt: FSCONV_num_(last[22]),
    closingCash: FSCONV_num_(last[23]),
    time: new Date().toISOString()
  }));
}

function FSCONV_readInfo_(sheet, label) {
  const target = FSCONV_key_(label);
  const values = sheet.getDataRange().getValues();
  for (let row = 0; row < values.length; row++) {
    if (FSCONV_key_(values[row][0]) === target) return values[row][1];
  }
  return '';
}

function FSCONV_num_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  const text = String(value == null ? '' : value).trim().replace(/\s/g, '');
  if (!text) return 0;
  const normalized = text.includes(',') && text.includes('.')
    ? text.replace(/\./g, '').replace(',', '.')
    : text.replace(/,/g, '');
  const number = Number(normalized);
  return isFinite(number) ? number : 0;
}

function FSCONV_norm_(value) {
  return String(value == null ? '' : value)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/\s+/g, ' ')
    .trim();
}

function FSCONV_key_(value) {
  return FSCONV_norm_(value).replace(/[^a-z0-9]/g, '');
}
