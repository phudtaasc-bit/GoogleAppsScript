const FS04A_CFG = Object.freeze({
  COST: '03. Chi phí & Vốn',
  CASH: '04. Dòng tiền & Tài trợ',
  SUMMARY: '04A. TH dòng tiền',
  OBSOLETE_SUMMARY: '04A. Tổng hợp dòng tiền',
  HEADER_ROW: 3,
  TOTAL_COLUMN: 3,
  FIRST_YEAR_COLUMN: 4,
  UNIT_DIVISOR: 1e9
});

function FS_lapSheet04A() {
  const ss = SpreadsheetApp.getActive();
  const costSheet = ss.getSheetByName(FS04A_CFG.COST);
  const cashSheet = ss.getSheetByName(FS04A_CFG.CASH);
  const summarySheet = ss.getSheetByName(FS04A_CFG.SUMMARY);

  if (!costSheet || !cashSheet) {
    throw new Error('Cần lập "03. Chi phí & Vốn" và "04. Dòng tiền & Tài trợ" trước.');
  }
  if (!summarySheet) {
    throw new Error('Không tìm thấy sheet mẫu "04A. TH dòng tiền". Code không tự tạo lại form báo cáo.');
  }

  const obsoleteSheet = ss.getSheetByName(FS04A_CFG.OBSOLETE_SUMMARY);
  if (obsoleteSheet && obsoleteSheet.getSheetId() !== summarySheet.getSheetId()) {
    ss.deleteSheet(obsoleteSheet);
  }

  const costByYear = FS04A_readCostByYear_(costSheet);
  const cashByYear = FS04A_readCashByYear_(cashSheet);
  const yearColumns = FS04A_readYearColumns_(summarySheet);

  if (!yearColumns.length) {
    throw new Error('Sheet "04A. TH dòng tiền" không có các cột năm tại hàng ' + FS04A_CFG.HEADER_ROW + '.');
  }

  const annualValues = {};
  yearColumns.forEach(item => {
    annualValues[item.year] = FS04A_buildYearValues_(
      costByYear[item.year] || FS04A_emptyCostYear_(),
      cashByYear[item.year] || FS04A_emptyCashYear_()
    );
  });

  const totalCost = FS04A_sumObjects_(Object.values(costByYear), FS04A_emptyCostYear_());
  const totalCash = FS04A_sumObjects_(Object.values(cashByYear), FS04A_emptyCashYear_());
  const totalValues = FS04A_buildYearValues_(totalCost, totalCash);

  FS04A_writeForm_(summarySheet, yearColumns, annualValues, totalValues);

  SpreadsheetApp.flush();
  return {
    years: yearColumns.map(item => item.year),
    annualValues,
    totalValues
  };
}

function FS04A_buildYearValues_(cost, cash) {
  const land = cost.landUse + cost.landRent;

  const totalInflow =
    cash.customerCash +
    cash.equityContribution +
    cash.loanDrawdown;

  const totalOutflow =
    cost.construction +
    cost.clearance +
    land +
    cost.infrastructure +
    cost.selling +
    cost.contingency +
    cost.operating +
    cost.maintenance +
    cost.vatIn +
    cash.vatPayable +
    cash.cit +
    cash.interest +
    cash.principalRepayment;

  return {
    5: cash.customerCash,
    6: cash.equityContribution,
    7: cash.loanDrawdown,
    8: totalInflow,

    10: cost.construction,
    11: cost.clearance,
    12: land,
    13: cost.infrastructure,
    14: cost.selling,
    15: cost.contingency,
    16: cost.operating,
    17: cost.maintenance,
    18: cost.vatIn,
    19: cash.vatPayable,
    20: cash.cit,
    21: cash.interest,
    22: cash.principalRepayment,
    23: totalOutflow,

    25: totalInflow - totalOutflow,
    26: cash.equityContribution,
    27: cash.loanDrawdown,
    28: cash.principalRepayment,
    29: cash.interest,

    // Chỉ tổng hợp trực tiếp từ Sheet 04, không tính lại tại 04A.
    30: cash.fcff,
    31: cash.fcfe
  };
}

function FS04A_writeForm_(sheet, yearColumns, annualValues, totalValues) {
  const rowNumbers = FS04A_outputRows_();
  const lastYearColumn = yearColumns[yearColumns.length - 1].column;
  const yearByColumn = {};
  yearColumns.forEach(item => {
    yearByColumn[item.column] = item.year;
  });

  rowNumbers.forEach(row => {
    const values = [FS04A_toReportUnit_(totalValues[row] || 0)];

    for (let column = FS04A_CFG.FIRST_YEAR_COLUMN; column <= lastYearColumn; column++) {
      const year = yearByColumn[column];
      const value = year ? (annualValues[year][row] || 0) : 0;
      values.push(FS04A_toReportUnit_(value));
    }

    sheet.getRange(row, FS04A_CFG.TOTAL_COLUMN, 1, values.length).setValues([values]);
  });
}

function FS04A_readCostByYear_(sheet) {
  const table = FS04A_readTable_(sheet);
  FS04A_require_(table.index, [
    'nam',
    'xdtbtruocvat',
    'gpmbtruocvat',
    'htkttruocvat',
    'tiensddtruocvat',
    'tienthuedattruocvat',
    'chiphibanhangtruocvat',
    'chiphivanhanhtruocvat',
    'chiphibaotritruocvat',
    'chiphiduphongtruocvat',
    'vatdauvao'
  ], sheet.getName());

  const result = {};
  table.values.forEach(row => {
    const year = FS04A_year_(row[table.index.nam]);
    if (!year) return;

    const item = result[year] || FS04A_emptyCostYear_();
    item.construction += FS04A_num_(row[table.index.xdtbtruocvat]);
    item.clearance += FS04A_num_(row[table.index.gpmbtruocvat]);
    item.infrastructure += FS04A_num_(row[table.index.htkttruocvat]);
    item.landUse += FS04A_num_(row[table.index.tiensddtruocvat]);
    item.landRent += FS04A_num_(row[table.index.tienthuedattruocvat]);
    item.selling += FS04A_num_(row[table.index.chiphibanhangtruocvat]);
    item.operating += FS04A_num_(row[table.index.chiphivanhanhtruocvat]);
    item.maintenance += FS04A_num_(row[table.index.chiphibaotritruocvat]);
    item.contingency += FS04A_num_(row[table.index.chiphiduphongtruocvat]);
    item.vatIn += FS04A_num_(row[table.index.vatdauvao]);
    result[year] = item;
  });

  return result;
}

function FS04A_readCashByYear_(sheet) {
  const table = FS04A_readTable_(sheet);
  FS04A_require_(table.index, [
    'nam',
    'dongtienkhachhang',
    'vatphainop',
    'thuetndn',
    'laivay',
    'vongopcsh',
    'giainganvay',
    'tragoc',
    'fcff',
    'fcfe'
  ], sheet.getName());

  const result = {};
  table.values.forEach(row => {
    const year = FS04A_year_(row[table.index.nam]);
    if (!year) return;

    const item = result[year] || FS04A_emptyCashYear_();
    item.customerCash += FS04A_num_(row[table.index.dongtienkhachhang]);
    item.vatPayable += FS04A_num_(row[table.index.vatphainop]);
    item.cit += FS04A_num_(row[table.index.thuetndn]);
    item.interest += FS04A_num_(row[table.index.laivay]);
    item.equityContribution += FS04A_num_(row[table.index.vongopcsh]);
    item.loanDrawdown += FS04A_num_(row[table.index.giainganvay]);
    item.principalRepayment += FS04A_num_(row[table.index.tragoc]);
    item.fcff += FS04A_num_(row[table.index.fcff]);
    item.fcfe += FS04A_num_(row[table.index.fcfe]);
    result[year] = item;
  });

  return result;
}

function FS04A_readYearColumns_(sheet) {
  const lastColumn = sheet.getLastColumn();
  const values = sheet.getRange(
    FS04A_CFG.HEADER_ROW,
    FS04A_CFG.FIRST_YEAR_COLUMN,
    1,
    Math.max(1, lastColumn - FS04A_CFG.FIRST_YEAR_COLUMN + 1)
  ).getValues()[0];

  const result = [];
  values.forEach((value, index) => {
    const year = FS04A_year_(value);
    if (year) {
      result.push({
        year,
        column: FS04A_CFG.FIRST_YEAR_COLUMN + index
      });
    }
  });

  return result;
}

function FS04A_outputRows_() {
  return [
    5, 6, 7, 8,
    10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23,
    25, 26, 27, 28, 29, 30, 31
  ];
}

function FS04A_sumObjects_(objects, seed) {
  const result = Object.assign({}, seed);
  objects.forEach(item => {
    Object.keys(result).forEach(key => {
      result[key] += FS04A_num_(item[key]);
    });
  });
  return result;
}

function FS04A_emptyCostYear_() {
  return {
    construction: 0,
    clearance: 0,
    infrastructure: 0,
    landUse: 0,
    landRent: 0,
    selling: 0,
    operating: 0,
    maintenance: 0,
    contingency: 0,
    vatIn: 0
  };
}

function FS04A_emptyCashYear_() {
  return {
    customerCash: 0,
    vatPayable: 0,
    cit: 0,
    interest: 0,
    equityContribution: 0,
    loanDrawdown: 0,
    principalRepayment: 0,
    fcff: 0,
    fcfe: 0
  };
}

function FS04A_readTable_(sheet) {
  const columns = sheet.getLastColumn();
  if (columns < 1 || sheet.getLastRow() < 1) {
    throw new Error('Sheet "' + sheet.getName() + '" không có dữ liệu.');
  }

  const headers = sheet.getRange(1, 1, 1, columns).getDisplayValues()[0];
  const values = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, columns).getValues()
    : [];

  const index = {};
  headers.forEach((header, position) => {
    index[FS04A_key_(header)] = position;
  });

  return { values, index };
}

function FS04A_require_(index, required, sheetName) {
  const missing = required.filter(key => index[key] == null);
  if (missing.length) {
    throw new Error('Sheet "' + sheetName + '" thiếu cột bắt buộc: ' + missing.join(', '));
  }
}

function FS04A_toReportUnit_(value) {
  return FS04A_num_(value) / FS04A_CFG.UNIT_DIVISOR;
}

function FS04A_year_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value.getFullYear();
  }

  if (typeof value === 'number' && isFinite(value)) {
    const year = Math.round(value);
    return year >= 1900 && year <= 3000 ? year : 0;
  }

  const digits = String(value == null ? '' : value).replace(/[^0-9]/g, '');
  if (digits.length < 4) return 0;

  const year = Number(digits.slice(-4));
  return year >= 1900 && year <= 3000 ? year : 0;
}

function FS04A_num_(value) {
  if (typeof value === 'number') {
    return isFinite(value) ? value : 0;
  }

  const text = String(value == null ? '' : value).trim().replace(/\s/g, '');
  if (!text) return 0;

  const normalized = text.includes(',') && text.includes('.')
    ? text.replace(/\./g, '').replace(',', '.')
    : text.replace(/,/g, '');

  const number = Number(normalized);
  return isFinite(number) ? number : 0;
}

function FS04A_norm_(value) {
  return String(value == null ? '' : value)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/\s+/g, ' ')
    .trim();
}

function FS04A_key_(value) {
  return FS04A_norm_(value)
    .replace(/²/g, '2')
    .replace(/\^2/g, '2')
    .replace(/m\s*2/g, 'm2')
    .replace(/[^a-z0-9]/g, '');
}
