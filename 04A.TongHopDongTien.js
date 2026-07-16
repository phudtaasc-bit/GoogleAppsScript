const FS04A_CFG = Object.freeze({
  REVENUE: '02. Doanh thu',
  COST: '03. Chi phí & Vốn',
  CASH: '04. Dòng tiền & Tài trợ',
  SUMMARY: '04A. Tổng hợp dòng tiền',
  TOLERANCE: 1
});

function FS_lapSheet04A() {
  const ss = SpreadsheetApp.getActive();
  const revenueSheet = ss.getSheetByName(FS04A_CFG.REVENUE);
  const costSheet = ss.getSheetByName(FS04A_CFG.COST);
  const cashSheet = ss.getSheetByName(FS04A_CFG.CASH);

  if (!revenueSheet || !costSheet || !cashSheet) {
    throw new Error('Cần lập "02. Doanh thu", "03. Chi phí & Vốn" và "04. Dòng tiền & Tài trợ" trước.');
  }

  const revenueByMonth = FS04A_readRevenue_(revenueSheet);
  const costByMonth = FS04A_readCost_(costSheet);
  const cashByMonth = FS04A_readCash_(cashSheet);
  const monthNumbers = Object.keys(cashByMonth).map(Number).sort((a, b) => a - b);

  if (!monthNumbers.length) throw new Error('Sheet "04. Dòng tiền & Tài trợ" không có dữ liệu.');

  const rows = monthNumbers.map(monthNo => {
    const revenue = revenueByMonth[monthNo] || FS04A_emptyRevenue_();
    const cost = costByMonth[monthNo] || FS04A_emptyCost_();
    const cash = cashByMonth[monthNo];

    const financingInflow = cash.equityContribution + cash.loanDrawdown;
    const totalCashInflow = cash.customerCash + financingInflow;

    const operatingOutflow = cash.costAfterVat + cash.vatPayable + cash.cit;
    const financingOutflow = cash.interest + cash.principalRepayment;
    const totalCashOutflow = operatingOutflow + financingOutflow;
    const netCashMovement = totalCashInflow - totalCashOutflow;

    const cashReconciliation = cash.openingCash + netCashMovement - cash.closingCash;
    const fcffCalculated = cash.customerCash - cash.costAfterVat - cash.vatPayable - cash.cit;
    const fcffDifference = cash.fcff - fcffCalculated;
    const fcfeCalculated = cash.closingCash - cash.equityContribution;
    const fcfeDifference = cash.fcfe - fcfeCalculated;
    const debtReconciliation = cash.openingDebt + cash.loanDrawdown - cash.principalRepayment - cash.closingDebt;

    const status = [cashReconciliation, fcffDifference, fcfeDifference, debtReconciliation]
      .every(value => Math.abs(value) <= FS04A_CFG.TOLERANCE)
      ? 'PASS'
      : 'FAIL';

    return [
      monthNo, cash.date, cash.year, cash.quarter,

      revenue.saleRevenue, revenue.rentRevenue, revenue.totalRevenue,
      cash.customerCash, cash.equityContribution, cash.loanDrawdown,
      financingInflow, totalCashInflow,

      cost.construction, cost.clearance, cost.infrastructure,
      cost.landUse, cost.landRent, cost.selling,
      cost.operating, cost.maintenance, cost.contingency,
      cash.costBeforeVat, cash.vatIn, cash.costAfterVat,
      cash.vatPayable, cash.cit, cash.interest, cash.principalRepayment,
      operatingOutflow, financingOutflow, totalCashOutflow,

      cash.openingCash, netCashMovement, cash.closingCash,
      cash.openingDebt, cash.closingDebt,
      cash.fcff, cash.fcfe,

      cashReconciliation, fcffDifference, fcfeDifference, debtReconciliation, status
    ];
  });

  FS04A_write_(ss, rows);

  const failed = rows.filter(row => row[row.length - 1] !== 'PASS');
  if (failed.length) {
    throw new Error(
      '04A có ' + failed.length + ' tháng đối chiếu FAIL. Kiểm tra các cột chênh lệch ở cuối sheet.'
    );
  }

  return rows;
}

function FS04A_readRevenue_(sheet) {
  const table = FS04A_readTable_(sheet);
  FS04A_require_(table.index, [
    'thangso', 'doanhthubantruocvat', 'doanhthuthuetruocvat', 'tongdoanhthutruocvat'
  ], sheet.getName());

  const result = {};
  table.values.forEach(row => {
    const monthNo = FS04A_num_(row[table.index.thangso]);
    if (monthNo < 1) return;
    const item = result[monthNo] || FS04A_emptyRevenue_();
    item.saleRevenue += FS04A_num_(row[table.index.doanhthubantruocvat]);
    item.rentRevenue += FS04A_num_(row[table.index.doanhthuthuetruocvat]);
    item.totalRevenue += FS04A_num_(row[table.index.tongdoanhthutruocvat]);
    result[monthNo] = item;
  });
  return result;
}

function FS04A_readCost_(sheet) {
  const table = FS04A_readTable_(sheet);
  FS04A_require_(table.index, [
    'thangso', 'xdtbtruocvat', 'gpmbtruocvat', 'htkttruocvat',
    'tiensddtruocvat', 'tienthuedattruocvat', 'chiphibanhangtruocvat',
    'chiphivanhanhtruocvat', 'chiphibaotritruocvat', 'chiphiduphongtruocvat'
  ], sheet.getName());

  const result = {};
  table.values.forEach(row => {
    const monthNo = FS04A_num_(row[table.index.thangso]);
    if (monthNo < 1) return;
    const item = result[monthNo] || FS04A_emptyCost_();
    item.construction += FS04A_num_(row[table.index.xdtbtruocvat]);
    item.clearance += FS04A_num_(row[table.index.gpmbtruocvat]);
    item.infrastructure += FS04A_num_(row[table.index.htkttruocvat]);
    item.landUse += FS04A_num_(row[table.index.tiensddtruocvat]);
    item.landRent += FS04A_num_(row[table.index.tienthuedattruocvat]);
    item.selling += FS04A_num_(row[table.index.chiphibanhangtruocvat]);
    item.operating += FS04A_num_(row[table.index.chiphivanhanhtruocvat]);
    item.maintenance += FS04A_num_(row[table.index.chiphibaotritruocvat]);
    item.contingency += FS04A_num_(row[table.index.chiphiduphongtruocvat]);
    result[monthNo] = item;
  });
  return result;
}

function FS04A_readCash_(sheet) {
  const table = FS04A_readTable_(sheet);
  FS04A_require_(table.index, [
    'thangso', 'thang', 'nam', 'quy', 'dongtienkhachhang',
    'tongchitruocvat', 'vatdauvao', 'tongchisauvat', 'vatphainop',
    'thuetndn', 'fcff', 'tientondauky', 'laivay', 'vongopcsh',
    'giainganvay', 'tragoc', 'dunodauky', 'dunocuoiky',
    'tientoncuoiky', 'fcfe'
  ], sheet.getName());

  const result = {};
  table.values.forEach(row => {
    const monthNo = FS04A_num_(row[table.index.thangso]);
    if (monthNo < 1) return;
    if (result[monthNo]) throw new Error('Sheet 04 bị trùng Tháng số: ' + monthNo);

    result[monthNo] = {
      date: row[table.index.thang],
      year: row[table.index.nam],
      quarter: row[table.index.quy],
      customerCash: FS04A_num_(row[table.index.dongtienkhachhang]),
      costBeforeVat: FS04A_num_(row[table.index.tongchitruocvat]),
      vatIn: FS04A_num_(row[table.index.vatdauvao]),
      costAfterVat: FS04A_num_(row[table.index.tongchisauvat]),
      vatPayable: FS04A_num_(row[table.index.vatphainop]),
      cit: FS04A_num_(row[table.index.thuetndn]),
      fcff: FS04A_num_(row[table.index.fcff]),
      openingCash: FS04A_num_(row[table.index.tientondauky]),
      interest: FS04A_num_(row[table.index.laivay]),
      equityContribution: FS04A_num_(row[table.index.vongopcsh]),
      loanDrawdown: FS04A_num_(row[table.index.giainganvay]),
      principalRepayment: FS04A_num_(row[table.index.tragoc]),
      openingDebt: FS04A_num_(row[table.index.dunodauky]),
      closingDebt: FS04A_num_(row[table.index.dunocuoiky]),
      closingCash: FS04A_num_(row[table.index.tientoncuoiky]),
      fcfe: FS04A_num_(row[table.index.fcfe])
    };
  });
  return result;
}

function FS04A_write_(ss, rows) {
  let sheet = ss.getSheetByName(FS04A_CFG.SUMMARY);
  if (!sheet) sheet = ss.insertSheet(FS04A_CFG.SUMMARY);
  sheet.clear();
  sheet.clearFormats();

  const headers = [[
    'Tháng số', 'Tháng', 'Năm', 'Quý',

    'Doanh thu bán trước VAT', 'Doanh thu thuê trước VAT', 'Tổng doanh thu trước VAT',
    'Dòng tiền khách hàng', 'Vốn góp CSH', 'Giải ngân vay',
    'Tổng dòng tiền tài trợ vào', 'Tổng dòng tiền vào',

    'XD/TB trước VAT', 'GPMB trước VAT', 'HTKT trước VAT',
    'Tiền SDĐ trước VAT', 'Tiền thuê đất trước VAT', 'Chi phí bán hàng trước VAT',
    'Chi phí vận hành trước VAT', 'Chi phí bảo trì trước VAT', 'Chi phí dự phòng trước VAT',
    'Tổng chi trước VAT', 'VAT đầu vào', 'Tổng chi sau VAT',
    'VAT phải nộp', 'Thuế TNDN', 'Lãi vay', 'Trả gốc',
    'Dòng tiền ra hoạt động/đầu tư', 'Dòng tiền ra tài trợ', 'Tổng dòng tiền ra',

    'Tiền tồn đầu kỳ', 'Dòng tiền thuần trong kỳ', 'Tiền tồn cuối kỳ',
    'Dư nợ đầu kỳ', 'Dư nợ cuối kỳ', 'FCFF', 'FCFE',

    'CL cân đối tiền', 'CL FCFF', 'CL FCFE', 'CL dư nợ', 'Trạng thái'
  ]];

  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
  if (rows.length) sheet.getRange(2, 1, rows.length, headers[0].length).setValues(rows);

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(4);
  sheet.getRange(1, 1, 1, headers[0].length)
    .setFontWeight('bold')
    .setWrap(true)
    .setVerticalAlignment('middle');

  sheet.getRange(1, 5, 1, 8).setBackground('#d9ead3');
  sheet.getRange(1, 13, 1, 19).setBackground('#f4cccc');
  sheet.getRange(1, 32, 1, 7).setBackground('#cfe2f3');
  sheet.getRange(1, 39, 1, 5).setBackground('#fff2cc');

  if (rows.length) {
    sheet.getRange(2, 2, rows.length, 1).setNumberFormat('MM/yyyy');
    sheet.getRange(2, 5, rows.length, 38).setNumberFormat('#,##0');
    sheet.getRange(2, 43, rows.length, 1).setHorizontalAlignment('center');
  }

  const widths = [
    70, 85, 65, 90,
    145, 145, 150, 145, 115, 115, 145, 135,
    115, 110, 110, 115, 120, 145, 145, 145, 145,
    125, 105, 125, 110, 110, 110, 105, 155, 135, 135,
    115, 135, 115, 110, 110, 115, 115,
    105, 90, 90, 90, 85
  ];
  widths.forEach((width, index) => sheet.setColumnWidth(index + 1, width));
  sheet.setRowHeight(1, 58);
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

function FS04A_emptyRevenue_() {
  return { saleRevenue: 0, rentRevenue: 0, totalRevenue: 0 };
}

function FS04A_emptyCost_() {
  return {
    construction: 0, clearance: 0, infrastructure: 0,
    landUse: 0, landRent: 0, selling: 0,
    operating: 0, maintenance: 0, contingency: 0
  };
}

function FS04A_num_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : 0;
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
