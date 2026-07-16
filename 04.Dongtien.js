const FS04_CFG = Object.freeze({
  TECH: '01A. Kỹ thuật',
  COST: '03. Chi phí & Vốn',
  PROFIT: '03A. Lợi nhuận & Thuế',
  CASH: '04. Dòng tiền & Tài trợ',
  CASH_LEGACY: '04. Dòng tiền'
});

function FS_lapSheet04() {
  return FS04_buildAndWrite_();
}

function FS_hoiTuTaiTro() {
  return FS04_buildAndWrite_();
}

function FS04_buildAndWrite_() {
  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName(FS04_CFG.TECH);
  const costSheet = ss.getSheetByName(FS04_CFG.COST);
  const profitSheet = ss.getSheetByName(FS04_CFG.PROFIT);

  if (!tech || !costSheet || !profitSheet) {
    throw new Error('Cần lập "01A. Kỹ thuật", "03. Chi phí & Vốn" và "03A. Lợi nhuận & Thuế" trước.');
  }

  const months = Math.max(0, FS04_num_(FS04_readInfoValue_(tech, 'Số tháng mô hình')));
  const loanRatio = FS04_rate_(FS04_readInfoValue_(tech, 'Tỷ lệ vốn vay'));
  const annualInterestRate = FS04_rate_(FS04_readInfoValue_(tech, 'Lãi suất vay năm'));

  if (!months) throw new Error('Số tháng mô hình phải lớn hơn 0.');
  if (loanRatio < 0 || loanRatio > 1) {
    throw new Error('Tỷ lệ vốn vay phải nằm trong khoảng 0% đến 100%.');
  }
  if (annualInterestRate < 0) throw new Error('Lãi suất vay năm không được âm.');

  const monthlyInterestRate = Math.pow(1 + annualInterestRate, 1 / 12) - 1;
  const costByMonth = FS04_readCostByMonth_(costSheet, months);
  const taxByMonth = FS04_readTaxByMonth_(profitSheet, months);
  const rows = [];

  let openingCash = 0;
  let openingDebt = 0;
  let openingVatCredit = 0;

  for (let monthNo = 1; monthNo <= months; monthNo++) {
    const current = costByMonth[monthNo] || FS04_emptyCostMonth_(monthNo);
    const corporateIncomeTax = taxByMonth[monthNo] || 0;

    const vatPayable = Math.max(0, current.vatOut - openingVatCredit - current.vatIn);
    const closingVatCredit = Math.max(0, openingVatCredit + current.vatIn - current.vatOut);

    const fcff = current.customerCash - current.costAfterVat - vatPayable - corporateIncomeTax;
    const interestExpense = openingDebt * monthlyInterestRate;

    // Tiền FCFE còn lại từ tháng trước được sử dụng trước khi huy động vốn mới.
    const cashBeforeFinancing = openingCash + fcff - interestExpense;

    let fundingNeed = 0;
    let equityContribution = 0;
    let loanDrawdown = 0;
    let principalRepayment = 0;
    let closingDebt = openingDebt;
    let closingCash = 0;

    if (cashBeforeFinancing < 0) {
      fundingNeed = -cashBeforeFinancing;
      loanDrawdown = fundingNeed * loanRatio;
      equityContribution = fundingNeed - loanDrawdown;
      closingDebt = openingDebt + loanDrawdown;
      closingCash = 0;
    } else {
      // Toàn bộ tiền dư sau chi phí và lãi vay được ưu tiên trả gốc.
      principalRepayment = Math.min(openingDebt, cashBeforeFinancing);
      closingDebt = Math.max(0, openingDebt - principalRepayment);

      // Phần còn lại là FCFE, tiếp tục nằm trong dự án và chuyển sang tháng sau.
      closingCash = Math.max(0, cashBeforeFinancing - principalRepayment);
    }

    const fcfe = closingCash;

    rows.push([
      monthNo,
      current.date,
      current.year,
      current.quarter,
      current.customerCash,
      current.vatOut,
      current.costBeforeVat,
      current.vatIn,
      current.costAfterVat,
      openingVatCredit,
      vatPayable,
      closingVatCredit,
      corporateIncomeTax,
      fcff,
      openingCash,
      cashBeforeFinancing,
      fundingNeed,
      interestExpense,
      equityContribution,
      loanDrawdown,
      principalRepayment,
      openingDebt,
      closingDebt,
      closingCash,
      fcfe
    ]);

    openingCash = closingCash;
    openingDebt = closingDebt;
    openingVatCredit = closingVatCredit;
  }

  FS04_write_(ss, rows);
  FS04_writeStatus_(rows, monthlyInterestRate, loanRatio);
  return rows;
}

function FS04_readCostByMonth_(sheet, months) {
  const table = FS04_readTable_(sheet);
  const required = [
    'thangso', 'thang', 'nam', 'quy',
    'dongtienkhachhang', 'vatdaura',
    'tongchitruocvat', 'vatdauvao', 'tongchisauvat'
  ];
  FS04_requireHeaders_(table.index, required, '03. Chi phí & Vốn');

  const result = {};
  for (let monthNo = 1; monthNo <= months; monthNo++) {
    result[monthNo] = FS04_emptyCostMonth_(monthNo);
  }

  table.values.forEach(row => {
    const monthNo = FS04_num_(row[table.index.thangso]);
    if (monthNo < 1 || monthNo > months) return;

    const item = result[monthNo];
    const date = row[table.index.thang];
    if (!item.date && date) item.date = date;
    if (!item.year) item.year = row[table.index.nam];
    if (!item.quarter) item.quarter = row[table.index.quy];

    item.customerCash += FS04_num_(row[table.index.dongtienkhachhang]);
    item.vatOut += FS04_num_(row[table.index.vatdaura]);
    item.costBeforeVat += FS04_num_(row[table.index.tongchitruocvat]);
    item.vatIn += FS04_num_(row[table.index.vatdauvao]);
    item.costAfterVat += FS04_num_(row[table.index.tongchisauvat]);
  });

  return result;
}

function FS04_readTaxByMonth_(sheet, months) {
  const table = FS04_readTable_(sheet);
  FS04_requireHeaders_(table.index, ['thangso', 'thuetndn'], '03A. Lợi nhuận & Thuế');

  const result = {};
  for (let monthNo = 1; monthNo <= months; monthNo++) result[monthNo] = 0;

  table.values.forEach(row => {
    const monthNo = FS04_num_(row[table.index.thangso]);
    if (monthNo < 1 || monthNo > months) return;
    result[monthNo] += FS04_num_(row[table.index.thuetndn]);
  });

  return result;
}

function FS04_readTable_(sheet) {
  if (sheet.getLastRow() < 1) {
    throw new Error('Sheet "' + sheet.getName() + '" không có dữ liệu.');
  }

  const columnCount = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, columnCount).getDisplayValues()[0];
  const values = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, columnCount).getValues()
    : [];
  const index = {};

  headers.forEach((header, position) => {
    index[FS04_key_(header)] = position;
  });

  return { headers, values, index };
}

function FS04_requireHeaders_(index, required, sheetName) {
  const missing = required.filter(key => index[key] == null);
  if (missing.length) {
    throw new Error('Sheet "' + sheetName + '" thiếu cột bắt buộc: ' + missing.join(', '));
  }
}

function FS04_emptyCostMonth_(monthNo) {
  return {
    monthNo,
    date: '',
    year: '',
    quarter: '',
    customerCash: 0,
    vatOut: 0,
    costBeforeVat: 0,
    vatIn: 0,
    costAfterVat: 0
  };
}

function FS04_write_(ss, rows) {
  let sheet = ss.getSheetByName(FS04_CFG.CASH);
  const legacy = ss.getSheetByName(FS04_CFG.CASH_LEGACY);

  if (!sheet && legacy) {
    legacy.setName(FS04_CFG.CASH);
    sheet = legacy;
  }
  if (!sheet) sheet = ss.insertSheet(FS04_CFG.CASH);

  sheet.clear();
  sheet.clearFormats();

  const headers = [[
    'Tháng số', 'Tháng', 'Năm', 'Quý',
    'Dòng tiền khách hàng', 'VAT đầu ra',
    'Tổng chi trước VAT', 'VAT đầu vào', 'Tổng chi sau VAT',
    'VAT khấu trừ đầu kỳ', 'VAT phải nộp', 'VAT khấu trừ cuối kỳ',
    'Thuế TNDN', 'FCFF',
    'FCFE đầu kỳ', 'Tiền trước tài trợ', 'Nhu cầu vốn',
    'Lãi vay', 'Vốn góp CSH', 'Giải ngân vay', 'Trả gốc',
    'Dư nợ đầu kỳ', 'Dư nợ cuối kỳ',
    'FCFE cuối kỳ', 'FCFE'
  ]];

  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, headers[0].length).setValues(rows);
  }

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(4);
  sheet.getRange(1, 1, 1, headers[0].length)
    .setFontWeight('bold')
    .setBackground('#ddebf7')
    .setWrap(true)
    .setVerticalAlignment('middle');

  if (rows.length) {
    sheet.getRange(2, 2, rows.length, 1).setNumberFormat('MM/yyyy');
    sheet.getRange(2, 5, rows.length, headers[0].length - 4).setNumberFormat('#,##0');
  }

  const widths = [
    70, 85, 65, 90, 145, 105, 135, 105, 135,
    120, 110, 120, 110, 120, 110, 130, 110, 110,
    115, 115, 105, 110, 110, 120, 120
  ];
  widths.forEach((width, index) => sheet.setColumnWidth(index + 1, width));
  sheet.setRowHeight(1, 48);
}

function FS04_writeStatus_(rows, monthlyRate, loanRatio) {
  const lastRow = rows.length ? rows[rows.length - 1] : [];
  const closingDebt = FS04_num_(lastRow[22]);
  const closingFcfe = FS04_num_(lastRow[23]);

  PropertiesService.getDocumentProperties().setProperty('FS_CONVERGENCE', JSON.stringify({
    converged: true,
    iterations: 1,
    maxDiff: 0,
    method: 'debt-sweep-fcfe-carry-forward',
    monthlyInterestRate: monthlyRate,
    loanRatio,
    closingDebt,
    closingFcfe,
    time: new Date().toISOString()
  }));
}

function FS04_readInfoValue_(sheet, label) {
  const target = FS04_key_(label);
  const values = sheet.getDataRange().getValues();

  for (let row = 0; row < values.length; row++) {
    if (FS04_key_(values[row][0]) === target) return values[row][1];
  }
  return '';
}

function FS04_num_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : 0;

  const text = String(value == null ? '' : value).trim().replace(/\s/g, '');
  if (!text) return 0;

  const normalized = text.includes(',') && text.includes('.')
    ? text.replace(/\./g, '').replace(',', '.')
    : text.replace(/,/g, '');
  const number = Number(normalized);
  return isFinite(number) ? number : 0;
}

function FS04_rate_(value) {
  if (typeof value === 'number') return value > 1 ? value / 100 : value;

  const text = String(value == null ? '' : value).trim();
  if (!text) return 0;

  const number = FS04_num_(text.replace('%', ''));
  return text.includes('%') || number > 1 ? number / 100 : number;
}

function FS04_norm_(value) {
  return String(value == null ? '' : value)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/\s+/g, ' ')
    .trim();
}

function FS04_key_(value) {
  return FS04_norm_(value)
    .replace(/²/g, '2')
    .replace(/\^2/g, '2')
    .replace(/m\s*2/g, 'm2')
    .replace(/[^a-z0-9]/g, '');
}
