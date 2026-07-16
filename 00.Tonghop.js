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

  FS00_require_(r.index, ['masp', 'tongdoanhthutruocvat', 'vatdaura'], revenue.getName());
  FS00_require_(c.index, [
    'xdtbtruocvat', 'gpmbtruocvat', 'htkttruocvat', 'tiensddtruocvat', 'tienthuedattruocvat',
    'chiphibanhangtruocvat', 'chiphivanhanhtruocvat', 'chiphibaotritruocvat',
    'chiphiduphongtruocvat', 'tongchisauvat'
  ], cost.getName());
  FS00_require_(p.index, ['lnst', 'thuetndn'], profit.getName());
  FS00_require_(f.index, [
    'thang', 'fcff', 'fcfe', 'vongopcsh', 'giainganvay', 'dunocuoiky', 'laivay', 'vatphainop'
  ], cash.getName());

  FS00_validateCashDates_(f, cash.getName());

  const projectName = FS00_readInfoValue_(tech, 'Tên dự án');
  const discountCell = FS00_findInfoCell_(tech, 'Tỷ suất chiết khấu');
  const loanRateCell = FS00_findInfoCell_(tech, 'Lãi suất vay năm');
  if (!discountCell) throw new Error('Không tìm thấy chỉ tiêu "Tỷ suất chiết khấu" tại 01A. Kỹ thuật.');
  if (!loanRateCell) throw new Error('Không tìm thấy chỉ tiêu "Lãi suất vay năm" tại 01A. Kỹ thuật.');

  const products = FS00_readProducts_(tech);
  if (!products.length) throw new Error('Block SAN_PHAM tại 01A. Kỹ thuật không có sản phẩm hợp lệ.');

  const revenueDetails = products.map(product => ({
    code: product.code,
    name: product.name,
    value: FS00_sumByCode_(r, product.code, ['tongdoanhthutruocvat', 'vatdaura'])
  }));

  const detailRows = FS00_syncRevenueRows_(summary, revenueDetails);
  const rows = FS00_summaryRows_(summary);

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
  if (equity + loan > totalInvestment + 1) {
    throw new Error(
      'Vốn CSH và vốn vay vượt Tổng vốn đầu tư. Tổng vốn đầu tư: ' +
      Math.round(totalInvestment).toLocaleString('vi-VN') +
      '; vốn CSH + vốn vay: ' + Math.round(equity + loan).toLocaleString('vi-VN') + ' đồng.'
    );
  }
  const customerFunding = Math.max(0, totalInvestment - equity - loan);
  const totalFunding = equity + loan + customerFunding;

  const totalRevenueWithVat = sumR('tongdoanhthutruocvat') + sumR('vatdaura');
  const pat = sumP('lnst');
  const cit = sumP('thuetndn');
  const vatPayable = sumF('vatphainop');
  const peakDebt = FS00_maxValue_(f, 'dunocuoiky');
  const fcff = FS00_columnValues_(f, 'fcff');
  const fcfe = FS00_columnValues_(f, 'fcfe');
  const paybackProject = FS00_payback_(fcff);
  const paybackEquity = FS00_payback_(fcfe);

  summary.getRange('A2:E2').breakApart();
  summary.getRange('A2:E2').merge().setValue(projectName ? 'DỰ ÁN: ' + projectName : 'DỰ ÁN');

  const billion = value => FS00_num_(value) / FS00_CFG.UNIT_DIVISOR;
  const coreInvestment = totalInvestment - selling - operating - maintenance;

  const valueMap = {
    C6: billion(construction), C7: billion(clearance), C8: billion(land),
    C9: billion(infrastructure), C10: billion(contingency), C11: billion(interest),
    C12: billion(totalInvestment), C13: billion(totalInvestmentExLand),
    C17: billion(equity), C18: billion(loan), C19: billion(customerFunding), C20: billion(totalFunding)
  };

  valueMap['D' + rows.totalRevenue] = billion(totalRevenueWithVat);
  detailRows.forEach((row, index) => {
    valueMap['D' + row] = billion(revenueDetails[index].value);
  });
  valueMap['D' + rows.coreInvestment] = billion(coreInvestment);
  valueMap['D' + rows.selling] = billion(selling);
  valueMap['D' + rows.operating] = billion(operating);
  valueMap['D' + rows.maintenance] = billion(maintenance);
  valueMap['D' + rows.pat] = billion(pat);
  valueMap['D' + rows.paybackProject] = paybackProject;
  valueMap['D' + rows.paybackEquity] = paybackEquity;
  valueMap['D' + rows.peakDebt] = billion(peakDebt);
  valueMap['D' + rows.interest] = billion(interest);
  valueMap['D' + rows.cit] = billion(cit);
  valueMap['D' + rows.vatPayable] = billion(vatPayable);
  Object.keys(valueMap).forEach(a1 => summary.getRange(a1).setValue(valueMap[a1]));

  summary.getRange('D' + rows.totalCost).setFormula(
    '=SUM(D' + rows.coreInvestment + ':D' + rows.maintenance + ')'
  );
  summary.getRange(rows.operating, 1, 1, 5).setFontWeight('normal');
  summary.getRange(rows.maintenance, 1, 1, 5).setFontWeight('normal');

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
  const dateCol = FS00_columnLetter_(f.index.thang + 1);
  const fcffCol = FS00_columnLetter_(f.index.fcff + 1);
  const fcfeCol = FS00_columnLetter_(f.index.fcfe + 1);
  const firstDataRow = 2;

  const dateAll = cashRef + '!' + dateCol + firstDataRow + ':' + dateCol + lastCashRow;
  const fcffAll = cashRef + '!' + fcffCol + firstDataRow + ':' + fcffCol + lastCashRow;
  const fcfeAll = cashRef + '!' + fcfeCol + firstDataRow + ':' + fcfeCol + lastCashRow;

  summary.getRange('D' + rows.npvProject).setFormula(
    '=IFERROR(XNPV($E$21;' + fcffAll + ';' + dateAll + ')/1E9;0)'
  );
  summary.getRange('D' + rows.irrProject).setFormula(
    '=IFERROR(XIRR(' + fcffAll + ';' + dateAll + ');0)'
  );
  summary.getRange('D' + rows.npvEquity).setFormula(
    '=IFERROR(XNPV($E$17;' + fcfeAll + ';' + dateAll + ')/1E9;0)'
  );
  summary.getRange('D' + rows.irrEquity).setFormula(
    '=IFERROR(XIRR(' + fcfeAll + ';' + dateAll + ');0)'
  );

  summary.getRange('C6:C20').setNumberFormat('#,##0.0');
  summary.getRange('D6:D20').setNumberFormat('0.0%');
  summary.getRange('E17:E21').setNumberFormat('0.00%');
  summary.getRange(rows.totalRevenue, 4, rows.pat - rows.totalRevenue + 1, 1).setNumberFormat('#,##0.0');
  summary.getRange('D' + rows.npvProject).setNumberFormat('#,##0.0');
  summary.getRange('D' + rows.irrProject).setNumberFormat('0.00%');
  summary.getRange('D' + rows.paybackProject).setNumberFormat('0.00');
  summary.getRange('D' + rows.npvEquity).setNumberFormat('#,##0.0');
  summary.getRange('D' + rows.irrEquity).setNumberFormat('0.00%');
  summary.getRange('D' + rows.paybackEquity).setNumberFormat('0.00');
  summary.getRange('D' + rows.peakDebt + ':D' + rows.vatPayable).setNumberFormat('#,##0.0');

  SpreadsheetApp.flush();
  return { rows, totalInvestment, totalFunding, peakDebt, productCount: products.length };
}

function FS00_syncRevenueRows_(sheet, revenueDetails) {
  const totalRevenueRow = FS00_findSummaryRow_(sheet, 'Tổng doanh thu có VAT');
  const totalCostRow = FS00_findSummaryRow_(sheet, 'Tổng chi phí có VAT');
  if (!totalRevenueRow || !totalCostRow || totalCostRow <= totalRevenueRow) {
    throw new Error('Không xác định được vùng Mục 1 và Mục 2 trên Sheet 00.');
  }

  const desiredCount = revenueDetails.length;
  const currentCount = totalCostRow - totalRevenueRow - 1;
  const formatSourceRow = currentCount > 0 ? totalRevenueRow + 1 : totalRevenueRow;

  if (desiredCount > currentCount) {
    const addCount = desiredCount - currentCount;
    sheet.insertRowsBefore(totalCostRow, addCount);
    sheet.getRange(formatSourceRow, 1, 1, 5).copyTo(
      sheet.getRange(totalCostRow, 1, addCount, 5),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false
    );
  } else if (desiredCount < currentCount) {
    sheet.deleteRows(totalRevenueRow + 1 + desiredCount, currentCount - desiredCount);
  }

  const rows = [];
  revenueDetails.forEach((item, index) => {
    const row = totalRevenueRow + 1 + index;
    rows.push(row);
    sheet.getRange(row, 1, 1, 5).clearContent();
    sheet.getRange(row, 1).setValue('1.' + (index + 1));
    sheet.getRange(row, 2).setValue('Phần ' + item.name);
    sheet.getRange(row, 3).setValue('tỷ đồng');
    sheet.getRange(row, 1, 1, 5).setFontWeight('normal');
  });

  return rows;
}

function FS00_readProducts_(sheet) {
  const block = FS00_findBlock_(sheet, 'SAN_PHAM');
  const values = sheet.getRange(block.startRow, 1, block.rowCount, 14).getValues();
  const seen = {};
  return values.map(row => ({
    code: String(row[0] || '').trim().toUpperCase(),
    name: String(row[1] || '').trim()
  })).filter(product => {
    if (!product.code || !product.name || seen[product.code]) return false;
    seen[product.code] = true;
    return true;
  });
}

function FS00_findBlock_(sheet, blockName) {
  const data = sheet.getDataRange().getDisplayValues();
  const target = FS00_key_(blockName);
  let blockRow = 0;
  for (let r = 0; r < data.length; r++) {
    if (data[r].some(value => FS00_key_(value) === target)) {
      blockRow = r + 1;
      break;
    }
  }
  if (!blockRow) throw new Error('Không tìm thấy block ' + blockName + ' tại ' + sheet.getName() + '.');

  const startRow = blockRow + 2;
  let endRow = startRow - 1;
  let blankCount = 0;
  const knownBlocks = ['thongtinchung', 'chiphi chung', 'chiphi chung vat dau vao van hanh', 'sanpham', 'kehoachbanthutien', 'tiendochiphi'];

  for (let row = startRow; row <= sheet.getLastRow(); row++) {
    const first = String(sheet.getRange(row, 1).getDisplayValue() || '').trim();
    const firstKey = FS00_key_(first);
    if (row > startRow && knownBlocks.some(name => firstKey === FS00_key_(name))) break;

    if (!first) {
      blankCount++;
      if (blankCount >= 2) break;
    } else {
      blankCount = 0;
      endRow = row;
    }
  }

  return {
    blockRow,
    startRow,
    endRow,
    rowCount: Math.max(0, endRow - startRow + 1)
  };
}

function FS00_ensureOperatingRows_(sheet) {
  const operatingRow = FS00_findSummaryRow_(sheet, 'Chi phí vận hành');
  const maintenanceRow = FS00_findSummaryRow_(sheet, 'Chi phí bảo trì');
  if (operatingRow && maintenanceRow) return;

  const patRow = FS00_findSummaryRow_(sheet, 'Lợi nhuận sau thuế');
  if (!patRow) throw new Error('Form Sheet 00 không có dòng "Lợi nhuận sau thuế".');

  sheet.insertRowsBefore(patRow, 2);
  sheet.getRange(patRow - 1, 1, 1, 5).copyTo(
    sheet.getRange(patRow, 1, 2, 5),
    SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
    false
  );
  sheet.getRange(patRow, 1).setValue('2.3');
  sheet.getRange(patRow, 2).setValue('Chi phí vận hành');
  sheet.getRange(patRow, 3).setValue('tỷ đồng');
  sheet.getRange(patRow + 1, 1).setValue('2.4');
  sheet.getRange(patRow + 1, 2).setValue('Chi phí bảo trì');
  sheet.getRange(patRow + 1, 3).setValue('tỷ đồng');
  sheet.getRange(patRow, 1, 2, 5).setFontWeight('normal');
}

function FS00_summaryRows_(sheet) {
  const aliases = {
    totalRevenue: ['Tổng doanh thu có VAT'],
    totalCost: ['Tổng chi phí có VAT'],
    coreInvestment: ['Tổng vốn đầu tư dự án'],
    selling: ['Chi phí bán hàng'],
    operating: ['Chi phí vận hành'],
    maintenance: ['Chi phí bảo trì'],
    pat: ['Lợi nhuận sau thuế'],
    npvProject: ['NPV dự án'],
    irrProject: ['IRR dự án'],
    paybackProject: ['Thời gian hoàn vốn dự án'],
    npvEquity: ['NPV vốn CSH'],
    irrEquity: ['IRR vốn CSH'],
    paybackEquity: ['Thời gian hoàn vốn - Vốn CSH', 'Thời gian hoàn vốn vốn CSH'],
    peakDebt: ['Đỉnh dư nợ vay'],
    interest: ['Tổng lãi vay'],
    cit: ['Tổng Thuế TNDN'],
    vatPayable: ['Tổng VAT phải nộp']
  };

  const rows = {};
  Object.keys(aliases).forEach(key => {
    rows[key] = FS00_findSummaryRowByAliases_(sheet, aliases[key]);
    if (!rows[key]) throw new Error('Không tìm thấy dòng "' + aliases[key][0] + '" trên Sheet 00.');
  });
  return rows;
}

function FS00_findSummaryRowByAliases_(sheet, aliases) {
  for (const label of aliases) {
    const row = FS00_findSummaryRow_(sheet, label);
    if (row) return row;
  }
  return 0;
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

function FS00_validateCashDates_(table, sheetName) {
  const position = table.index.thang;
  for (let i = 0; i < table.values.length; i++) {
    const value = table.values[i][position];
    if (!(value instanceof Date) || isNaN(value.getTime())) {
      throw new Error(
        'Cột "Tháng" tại sheet "' + sheetName + '" phải là ngày thực. Lỗi tại dòng ' + (i + 2) + '.'
      );
    }
  }
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

function FS00_maxValue_(table, key) {
  const values = FS00_columnValues_(table, key);
  return values.length ? Math.max.apply(null, values) : 0;
}

function FS00_sumByCode_(table, code, keys) {
  const codePosition = table.index.masp;
  const normalizedCode = String(code || '').trim().toUpperCase();
  return table.values.reduce((sum, row) => {
    const rowCode = String(row[codePosition] || '').trim().toUpperCase();
    if (rowCode !== normalizedCode) return sum;
    return sum + keys.reduce((subtotal, key) => {
      const position = table.index[key];
      return subtotal + (position == null ? 0 : FS00_num_(row[position]));
    }, 0);
  }, 0);
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
