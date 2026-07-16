const FS00_CFG = Object.freeze({
  TECH: '01A. Kỹ thuật',
  REVENUE: '02. Doanh thu',
  COST: '03. Chi phí & Vốn',
  PROFIT: '03A. Lợi nhuận & Thuế',
  CASH: '04. Dòng tiền & Tài trợ',
  SUMMARY: '00. Tổng hợp',
  UNIT_DIVISOR: 1e9
});

function FS_lapSheet00() {
  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName(FS00_CFG.TECH);
  const revenue = ss.getSheetByName(FS00_CFG.REVENUE);
  const cost = ss.getSheetByName(FS00_CFG.COST);
  const profit = ss.getSheetByName(FS00_CFG.PROFIT);
  const cash = ss.getSheetByName(FS00_CFG.CASH);
  const summary = ss.getSheetByName(FS00_CFG.SUMMARY);

  if (!tech || !revenue || !cost || !profit || !cash) {
    throw new Error('Cần lập đủ các sheet 01A, 02, 03, 03A và 04 trước khi tổng hợp Sheet 00.');
  }
  if (!summary) throw new Error('Không tìm thấy sheet mẫu "00. Tổng hợp".');

  FS00_ensureOperatingRows_(summary);

  const r = FS00_readTable_(revenue);
  const c = FS00_readTable_(cost);
  const p = FS00_readTable_(profit);
  const f = FS00_readTable_(cash);

  FS00_require_(r.index, [
    'masp','tongdoanhthutruocvat','vatdaura'
  ], revenue.getName());
  FS00_require_(c.index, [
    'xdtbtruocvat','gpmbtruocvat','htkttruocvat','tiensddtruocvat','tienthuedattruocvat',
    'chiphibanhangtruocvat','chiphivanhanhtruocvat','chiphibaotritruocvat',
    'chiphiduphongtruocvat','tongchisauvat'
  ], cost.getName());
  FS00_require_(p.index, ['lnst','thuetndn'], profit.getName());
  FS00_require_(f.index, [
    'fcff','fcfe','vongopcsh','giainganvay','dunocuoiky','laivay','vatphainop'
  ], cash.getName());

  const projectName = FS00_readInfoValue_(tech, 'Tên dự án');
  const discountCell = FS00_findInfoCell_(tech, 'Tỷ suất chiết khấu');
  const loanRateCell = FS00_findInfoCell_(tech, 'Lãi suất vay năm');
  if (!discountCell) throw new Error('Không tìm thấy chỉ tiêu "Tỷ suất chiết khấu" tại 01A. Kỹ thuật.');
  if (!loanRateCell) throw new Error('Không tìm thấy chỉ tiêu "Lãi suất vay năm" tại 01A. Kỹ thuật.');

  const sumR = key => FS00_sumColumn_(r, key);
  const sumC = key => FS00_sumColumn_(c, key);
  const sumP = key => FS00_sumColumn_(p, key);
  const sumF = key => FS00_sumColumn_(f, key);

  const construction = sumC('xdtbtruocvat');
  const clearance = sumC('gpmbtruocvat');
  const land = sumC('tiensddtruocvat') + sumC('tienthuedattruocvat');
  const infrastructure = sumC('htkttruocvat');
  const contingency = sumC('chiphiduphongtruocvat');
  const selling = sumC('chiphibanhangtruocvat');
  const operating = sumC('chiphivanhanhtruocvat');
  const maintenance = sumC('chiphibaotritruocvat');
  const interest = sumF('laivay');
  const totalCostAfterVat = sumC('tongchisauvat');
  const totalInvestment = totalCostAfterVat + interest;
  const totalInvestmentExLand = totalInvestment - land;

  const equity = sumF('vongopcsh');
  const loan = sumF('giainganvay');
  const customerFunding = sumR('tongdoanhthutruocvat') + sumR('vatdaura');
  const totalFunding = equity + loan + customerFunding;

  const totalRevenueWithVat = customerFunding;
  const revenueCC = FS00_sumByCode_(r, 'CC', ['tongdoanhthutruocvat','vatdaura']);
  const revenueLK = FS00_sumByCode_(r, 'LK', ['tongdoanhthutruocvat','vatdaura']);
  const revenueRent =
    FS00_sumByCode_(r, 'TMDV', ['tongdoanhthutruocvat','vatdaura']) +
    FS00_sumByCode_(r, 'CHO', ['tongdoanhthutruocvat','vatdaura']);

  const pat = sumP('lnst');
  const cit = sumP('thuetndn');
  const vatPayable = sumF('vatphainop');
  const endingDebt = FS00_lastValue_(f, 'dunocuoiky');
  const fcff = FS00_columnValues_(f, 'fcff');
  const fcfe = FS00_columnValues_(f, 'fcfe');
  const paybackProject = FS00_payback_(fcff);
  const paybackEquity = FS00_payback_(fcfe);

  summary.getRange('A2:E2').breakApart();
  summary.getRange('A2:E2').merge().setValue(projectName ? 'DỰ ÁN: ' + projectName : 'DỰ ÁN');

  summary.getRange('B26').setValue('Phần Chung cư');
  summary.getRange('B27').setValue('Phần Liền kề');
  summary.getRange('B28').setValue('Phần TMDV / Chợ cho thuê');

  const rows = FS00_summaryRows_(summary);
  const billion = value => FS00_num_(value) / FS00_CFG.UNIT_DIVISOR;

  const valueMap = {
    C6: billion(construction), C7: billion(clearance), C8: billion(land),
    C9: billion(infrastructure), C10: billion(contingency), C11: billion(interest),
    C12: billion(totalInvestment), C13: billion(totalInvestmentExLand),
    C17: billion(equity), C18: billion(loan), C19: billion(customerFunding), C20: billion(totalFunding),
    D25: billion(totalRevenueWithVat), D26: billion(revenueCC), D27: billion(revenueLK), D28: billion(revenueRent),
    D29: billion(totalCostAfterVat), D30: billion(totalInvestment), D31: billion(selling),
    D39: billion(endingDebt), D40: billion(interest), D41: billion(cit), D42: billion(vatPayable)
  };

  valueMap['D' + rows.operating] = billion(operating);
  valueMap['D' + rows.maintenance] = billion(maintenance);
  valueMap['D' + rows.pat] = billion(pat);
  valueMap['D' + rows.paybackProject] = paybackProject;
  valueMap['D' + rows.paybackEquity] = paybackEquity;
  valueMap['D' + rows.endingDebt] = billion(endingDebt);
  valueMap['D' + rows.interest] = billion(interest);
  valueMap['D' + rows.cit] = billion(cit);
  valueMap['D' + rows.vatPayable] = billion(vatPayable);

  Object.keys(valueMap).forEach(a1 => summary.getRange(a1).setValue(valueMap[a1]));

  const ratioMap = {
    D6: totalInvestment ? construction / totalInvestment : 0,
    D7: totalInvestment ? clearance / totalInvestment : 0,
    D8: totalInvestment ? land / totalInvestment : 0,
    D9: totalInvestment ? infrastructure / totalInvestment : 0,
    D10: totalInvestment ? contingency / totalInvestment : 0,
    D11: totalInvestment ? interest / totalInvestment : 0,
    D12: totalInvestment ? 1 : 0,
    D17: totalFunding ? equity / totalFunding : 0,
    D18: totalFunding ? loan / totalFunding : 0,
    D19: totalFunding ? customerFunding / totalFunding : 0,
    D20: totalFunding ? 1 : 0
  };
  Object.keys(ratioMap).forEach(a1 => summary.getRange(a1).setValue(ratioMap[a1]));

  const techSheetRef = FS00_quoteSheet_(tech.getName());
  summary.getRange('E17').setFormula('=' + techSheetRef + '!' + discountCell.getA1Notation());
  summary.getRange('E18').setFormula('=' + techSheetRef + '!' + loanRateCell.getA1Notation());
  summary.getRange('E19').setValue(0);
  summary.getRange('E21').setFormula('=D17*E17+D18*E18+D19*E19');

  const cashRef = FS00_quoteSheet_(cash.getName());
  const lastCashRow = cash.getLastRow();
  const fcffCol = FS00_columnLetter_(f.index.fcff + 1);
  const fcfeCol = FS00_columnLetter_(f.index.fcfe + 1);
  const firstDataRow = 2;
  const secondDataRow = Math.min(3, lastCashRow);

  const fcffAll = cashRef + '!' + fcffCol + firstDataRow + ':' + fcffCol + lastCashRow;
  const fcfeAll = cashRef + '!' + fcfeCol + firstDataRow + ':' + fcfeCol + lastCashRow;
  const fcffAfterFirst = cashRef + '!' + fcffCol + secondDataRow + ':' + fcffCol + lastCashRow;
  const fcfeAfterFirst = cashRef + '!' + fcfeCol + secondDataRow + ':' + fcfeCol + lastCashRow;
  const fcffFirst = cashRef + '!' + fcffCol + firstDataRow;
  const fcfeFirst = cashRef + '!' + fcfeCol + firstDataRow;

  summary.getRange('D' + rows.npvProject).setFormula(
    '=IFERROR((NPV((1+$E$21)^(1/12)-1,' + fcffAfterFirst + ')+' + fcffFirst + ')/1E9,0)'
  );
  summary.getRange('D' + rows.irrProject).setFormula(
    '=IFERROR((1+IRR(' + fcffAll + '))^12-1,0)'
  );
  summary.getRange('D' + rows.npvEquity).setFormula(
    '=IFERROR((NPV((1+$E$17)^(1/12)-1,' + fcfeAfterFirst + ')+' + fcfeFirst + ')/1E9,0)'
  );
  summary.getRange('D' + rows.irrEquity).setFormula(
    '=IFERROR((1+IRR(' + fcfeAll + '))^12-1,0)'
  );

  summary.getRange('C6:C20').setNumberFormat('#,##0.0');
  summary.getRange('D6:D20').setNumberFormat('0.0%');
  summary.getRange('E17:E21').setNumberFormat('0.00%');
  summary.getRange('D25:D' + rows.pat).setNumberFormat('#,##0.0');
  summary.getRange('D' + rows.npvProject).setNumberFormat('#,##0.0');
  summary.getRange('D' + rows.irrProject).setNumberFormat('0.00%');
  summary.getRange('D' + rows.paybackProject).setNumberFormat('0.00');
  summary.getRange('D' + rows.npvEquity).setNumberFormat('#,##0.0');
  summary.getRange('D' + rows.irrEquity).setNumberFormat('0.00%');
  summary.getRange('D' + rows.paybackEquity).setNumberFormat('0.00');
  summary.getRange('D' + rows.endingDebt + ':D' + rows.vatPayable).setNumberFormat('#,##0.0');

  SpreadsheetApp.flush();
  return { rows, totalInvestment, totalFunding };
}

function FS00_ensureOperatingRows_(sheet) {
  const operatingRow = FS00_findSummaryRow_(sheet, 'Chi phí vận hành');
  const maintenanceRow = FS00_findSummaryRow_(sheet, 'Chi phí bảo trì');
  if (operatingRow && maintenanceRow) return;

  const patRow = FS00_findSummaryRow_(sheet, 'Lợi nhuận sau thuế');
  if (!patRow) throw new Error('Form Sheet 00 không có dòng "Lợi nhuận sau thuế".');

  sheet.insertRowsBefore(patRow, 2);
  sheet.getRange(patRow, 1, 2, 5).setBorder(true, true, true, true, true, true);
  sheet.getRange(patRow, 1, 2, 5).setBackground('#ffffff').setFontColor('#000000');
  sheet.getRange(patRow, 1).setValue('2.3');
  sheet.getRange(patRow, 2).setValue('Chi phí vận hành');
  sheet.getRange(patRow, 3).setValue('tỷ đồng');
  sheet.getRange(patRow + 1, 1).setValue('2.4');
  sheet.getRange(patRow + 1, 2).setValue('Chi phí bảo trì');
  sheet.getRange(patRow + 1, 3).setValue('tỷ đồng');
}

function FS00_summaryRows_(sheet) {
  const required = {
    operating: 'Chi phí vận hành', maintenance: 'Chi phí bảo trì', pat: 'Lợi nhuận sau thuế',
    npvProject: 'NPV dự án', irrProject: 'IRR dự án', paybackProject: 'Thời gian hoàn vốn dự án',
    npvEquity: 'NPV vốn CSH', irrEquity: 'IRR vốn CSH', paybackEquity: 'Thời gian hoàn vốn - Vốn CSH',
    endingDebt: 'Dư nợ cuối vay', interest: 'Tổng lãi vay', cit: 'Tổng Thuế TNDN', vatPayable: 'Tổng VAT phải nộp'
  };
  const rows = {};
  Object.keys(required).forEach(key => {
    rows[key] = FS00_findSummaryRow_(sheet, required[key]);
    if (!rows[key]) throw new Error('Không tìm thấy dòng "' + required[key] + '" trên Sheet 00.');
  });
  return rows;
}

function FS00_findSummaryRow_(sheet, label) {
  const target = FS00_key_(label);
  const values = sheet.getRange(1, 2, sheet.getLastRow(), 1).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    if (FS00_key_(values[i][0]) === target) return i + 1;
  }
  return 0;
}

function FS00_readTable_(sheet) {
  const columns = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, columns).getDisplayValues()[0];
  const values = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, columns).getValues()
    : [];
  const index = {};
  headers.forEach((header, position) => { index[FS00_key_(header)] = position; });
  return { values, index };
}

function FS00_require_(index, required, sheetName) {
  const missing = required.filter(key => index[key] == null);
  if (missing.length) throw new Error('Sheet "' + sheetName + '" thiếu cột: ' + missing.join(', '));
}

function FS00_sumColumn_(table, key) {
  const position = table.index[key];
  if (position == null) return 0;
  return table.values.reduce((sum, row) => sum + FS00_num_(row[position]), 0);
}

function FS00_columnValues_(table, key) {
  const position = table.index[key];
  if (position == null) return [];
  return table.values.map(row => FS00_num_(row[position]));
}

function FS00_sumByCode_(table, code, keys) {
  const codePosition = table.index.masp;
  return table.values.reduce((sum, row) => {
    const rowCode = String(row[codePosition] || '').trim().toUpperCase();
    if (rowCode !== code) return sum;
    return sum + keys.reduce((subtotal, key) => {
      const position = table.index[key];
      return subtotal + (position == null ? 0 : FS00_num_(row[position]));
    }, 0);
  }, 0);
}

function FS00_lastValue_(table, key) {
  const position = table.index[key];
  if (position == null || !table.values.length) return 0;
  return FS00_num_(table.values[table.values.length - 1][position]);
}

function FS00_readInfoValue_(sheet, label) {
  const cell = FS00_findInfoCell_(sheet, label);
  return cell ? cell.getValue() : '';
}

function FS00_findInfoCell_(sheet, label) {
  const target = FS00_key_(label);
  const values = sheet.getDataRange().getDisplayValues();
  for (let row = 0; row < values.length; row++) {
    if (FS00_key_(values[row][0]) === target) return sheet.getRange(row + 1, 2);
  }
  return null;
}

function FS00_payback_(cashFlows) {
  let cumulative = 0;
  for (let index = 0; index < cashFlows.length; index++) {
    const current = FS00_num_(cashFlows[index]);
    const previous = cumulative;
    cumulative += current;
    if (cumulative >= 0 && previous < 0 && current !== 0) {
      return index + Math.abs(previous) / current;
    }
  }
  return 0;
}

function FS00_quoteSheet_(name) {
  return "'" + String(name).replace(/'/g, "''") + "'";
}

function FS00_columnLetter_(column) {
  let result = '';
  let value = column;
  while (value > 0) {
    value--;
    result = String.fromCharCode(65 + (value % 26)) + result;
    value = Math.floor(value / 26);
  }
  return result;
}

function FS00_num_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  const text = String(value == null ? '' : value).trim().replace(/\s/g, '');
  if (!text) return 0;
  const normalized = text.includes(',') && text.includes('.')
    ? text.replace(/\./g, '').replace(',', '.')
    : text.replace(/,/g, '');
  const number = Number(normalized);
  return isFinite(number) ? number : 0;
}

function FS00_norm_(value) {
  return String(value == null ? '' : value)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/\s+/g, ' ')
    .trim();
}

function FS00_key_(value) {
  return FS00_norm_(value)
    .replace(/²/g, '2')
    .replace(/\^2/g, '2')
    .replace(/m\s*2/g, 'm2')
    .replace(/[^a-z0-9]/g, '');
}
