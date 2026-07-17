function FS_lapSheet00_TheoDanhMucVaSanPham() {
  const result = FS_lapSheet00_TheoDanhMuc();
  FS00P_capNhatCotSanPham_();
  return result;
}

function FS00P_capNhatCotSanPham_() {
  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName('01A. Kỹ thuật');
  const revenueSheet = ss.getSheetByName('02. Doanh thu');
  const costSheet = ss.getSheetByName('03. Chi phí & Vốn');
  const profitSheet = ss.getSheetByName('03A. Lợi nhuận & Thuế');
  const summary = ss.getSheetByName('00. Tổng hợp');

  if (!tech || !revenueSheet || !costSheet || !profitSheet || !summary) {
    throw new Error('Thiếu sheet nguồn để tổng hợp chỉ tiêu theo sản phẩm.');
  }

  const products = FS00P_docSanPham_(tech);
  if (!products.length) throw new Error('Block SAN_PHAM không có sản phẩm hợp lệ.');

  const revenue = FS00_readTable_(revenueSheet);
  const cost = FS00_readTable_(costSheet);
  const profit = FS00_readTable_(profitSheet);

  FS00_require_(revenue.index, ['masp', 'tongdoanhthutruocvat', 'vatdaura'], revenueSheet.getName());
  FS00_require_(cost.index, [
    'masp', 'chiphibanhangtruocvat', 'chiphivanhanhtruocvat',
    'chiphibaotritruocvat', 'tongchisauvat'
  ], costSheet.getName());
  FS00_require_(profit.index, ['masp', 'lnst', 'thuetndn'], profitSheet.getName());

  const firstProductCol = 6; // F
  const oldLastCol = Math.max(summary.getLastColumn(), firstProductCol);
  const requiredLastCol = firstProductCol + products.length - 1;

  if (summary.getMaxColumns() < requiredLastCol) {
    summary.insertColumnsAfter(summary.getMaxColumns(), requiredLastCol - summary.getMaxColumns());
  }

  if (oldLastCol >= firstProductCol) {
    summary.getRange(23, firstProductCol, summary.getMaxRows() - 22, oldLastCol - firstProductCol + 1)
      .clearContent()
      .clearFormat();
  }

  const rows = FS00P_docDongChiTieu_(summary);
  const lastIndicatorRow = rows.vatPayable;

  summary.getRange('A23:E23').breakApart();
  summary.getRange(23, 1, 1, requiredLastCol).breakApart().merge()
    .setValue('III. ĐÁNH GIÁ HIỆU QUẢ ĐẦU TƯ');

  const baseHeader = summary.getRange(24, 4);
  const headerRange = summary.getRange(24, firstProductCol, 1, products.length);
  baseHeader.copyTo(headerRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  headerRange.setValues([products.map(p => p.name + (p.group ? ' - ' + p.group : ''))]);
  headerRange
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  const vatRates = FS00_docVatRates_(tech);
  const detailRows = [];
  for (let row = rows.totalRevenue + 1; row < rows.totalCost; row++) detailRows.push(row);

  products.forEach((product, productIndex) => {
    const col = firstProductCol + productIndex;
    const code = product.code;

    const revenueWithVat = FS00P_sumTheoMa_(revenue, code, ['tongdoanhthutruocvat', 'vatdaura']);
    const totalCostAfterVat = FS00P_sumTheoMa_(cost, code, ['tongchisauvat']);
    const selling = FS00P_sumTheoMa_(cost, code, ['chiphibanhangtruocvat']) * (1 + vatRates.selling);
    const operating = FS00P_sumTheoMa_(cost, code, ['chiphivanhanhtruocvat']) * (1 + vatRates.operating);
    const maintenance = FS00P_sumTheoMa_(cost, code, ['chiphibaotritruocvat']) * (1 + vatRates.maintenance);
    const allocatedInterest = FS00P_sumTheoMaAliases_(profit, code, [
      'laivayvonhoaphanbo', 'laivayphanbo', 'laivay'
    ]);
    const coreInvestment = Math.max(0, totalCostAfterVat - selling - operating - maintenance + allocatedInterest);
    const totalCost = coreInvestment + selling + operating + maintenance;
    const pat = FS00P_sumTheoMa_(profit, code, ['lnst']);
    const cit = FS00P_sumTheoMa_(profit, code, ['thuetndn']);
    const vatPayable = FS00P_sumTheoMaAliases_(profit, code, ['vatphainop']);

    const fcff = FS00P_chuoiTheoMaAliases_(profit, code, ['fcff']);
    const fcfe = FS00P_chuoiTheoMaAliases_(profit, code, ['fcfe']);
    const dates = FS00P_chuoiNgayTheoMa_(profit, code);
    const projectNpv = fcff.length && dates.length === fcff.length
      ? FS00P_xnpv_(FS00_num_(summary.getRange('E21').getValue()), fcff, dates) / 1e9
      : 0;
    const projectIrr = fcff.length && dates.length === fcff.length ? FS00P_xirr_(fcff, dates) : 0;
    const projectPayback = fcff.length ? FS00_tinhHoanVonBenVung_(fcff) : 0;
    const equityNpv = fcfe.length && dates.length === fcfe.length
      ? FS00P_xnpv_(FS00_num_(summary.getRange('E17').getValue()), fcfe, dates) / 1e9
      : 0;
    const equityIrr = fcfe.length && dates.length === fcfe.length ? FS00P_xirr_(fcfe, dates) : 0;
    const equityPayback = fcfe.length ? FS00_tinhHoanVonBenVung_(fcfe) : 0;
    const peakDebt = FS00P_maxTheoMaAliases_(profit, code, ['dunocuoiky']);

    const values = {};
    values[rows.totalRevenue] = revenueWithVat / 1e9;
    detailRows.forEach((row, index) => {
      values[row] = index === productIndex ? revenueWithVat / 1e9 : 0;
    });
    values[rows.totalCost] = totalCost / 1e9;
    values[rows.coreInvestment] = coreInvestment / 1e9;
    values[rows.selling] = selling / 1e9;
    values[rows.operating] = operating / 1e9;
    values[rows.maintenance] = maintenance / 1e9;
    values[rows.pat] = pat / 1e9;
    values[rows.npvProject] = projectNpv;
    values[rows.irrProject] = projectIrr;
    values[rows.paybackProject] = projectPayback;
    values[rows.npvEquity] = equityNpv;
    values[rows.irrEquity] = equityIrr;
    values[rows.paybackEquity] = equityPayback;
    values[rows.peakDebt] = peakDebt / 1e9;
    values[rows.interest] = allocatedInterest / 1e9;
    values[rows.cit] = cit / 1e9;
    values[rows.vatPayable] = vatPayable / 1e9;

    const output = [];
    for (let row = rows.totalRevenue; row <= lastIndicatorRow; row++) {
      output.push([Object.prototype.hasOwnProperty.call(values, row) ? FS00_num_(values[row]) : 0]);
    }

    const outputRange = summary.getRange(rows.totalRevenue, col, output.length, 1);
    summary.getRange(rows.totalRevenue, 4, output.length, 1)
      .copyTo(outputRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    outputRange.setValues(output);

    summary.getRange(rows.totalRevenue, col, rows.pat - rows.totalRevenue + 1, 1)
      .setNumberFormat('#,##0.0');
    summary.getRange(rows.npvProject, col).setNumberFormat('#,##0.0');
    summary.getRange(rows.irrProject, col).setNumberFormat('0.00%');
    summary.getRange(rows.paybackProject, col).setNumberFormat('0.00');
    summary.getRange(rows.npvEquity, col).setNumberFormat('#,##0.0');
    summary.getRange(rows.irrEquity, col).setNumberFormat('0.00%');
    summary.getRange(rows.paybackEquity, col).setNumberFormat('0.00');
    summary.getRange(rows.peakDebt, col, rows.vatPayable - rows.peakDebt + 1, 1)
      .setNumberFormat('#,##0.0');
  });

  summary.getRange(23, 1, lastIndicatorRow - 22, requiredLastCol)
    .setFontFamily('Times New Roman')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  for (let col = firstProductCol; col <= requiredLastCol; col++) summary.setColumnWidth(col, 125);
  SpreadsheetApp.flush();
}

function FS00P_docSanPham_(sheet) {
  const block = FS00_findBlock_(sheet, 'SAN_PHAM');
  const values = sheet.getRange(block.startRow, 1, block.rowCount, 14).getValues();
  const products = [];
  const seen = {};

  values.forEach(row => {
    const code = String(row[0] || '').trim().toUpperCase();
    const name = String(row[1] || '').trim();
    const group = String(row[2] || '').trim();
    if (!code || !name || seen[code]) return;
    seen[code] = true;
    products.push({ code, name, group });
  });
  return products;
}

function FS00P_docDongChiTieu_(sheet) {
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

function FS00P_sumTheoMa_(table, code, keys) {
  return FS00P_sumTheoMaAliases_(table, code, keys);
}

function FS00P_sumTheoMaAliases_(table, code, aliases) {
  const codePos = table.index.masp;
  if (codePos == null) return 0;
  const positions = aliases.map(alias => table.index[FS00_key_(alias)]).filter(pos => pos != null);
  if (!positions.length) return 0;
  const normalizedCode = String(code || '').trim().toUpperCase();
  return table.values.reduce((sum, row) => {
    if (String(row[codePos] || '').trim().toUpperCase() !== normalizedCode) return sum;
    return sum + positions.reduce((subtotal, pos) => subtotal + FS00_num_(row[pos]), 0);
  }, 0);
}

function FS00P_chuoiTheoMaAliases_(table, code, aliases) {
  const codePos = table.index.masp;
  if (codePos == null) return [];
  let valuePos = null;
  for (const alias of aliases) {
    const pos = table.index[FS00_key_(alias)];
    if (pos != null) { valuePos = pos; break; }
  }
  if (valuePos == null) return [];
  const normalizedCode = String(code || '').trim().toUpperCase();
  return table.values
    .filter(row => String(row[codePos] || '').trim().toUpperCase() === normalizedCode)
    .map(row => FS00_num_(row[valuePos]));
}

function FS00P_chuoiNgayTheoMa_(table, code) {
  const codePos = table.index.masp;
  const datePos = table.index.thang;
  if (codePos == null || datePos == null) return [];
  const normalizedCode = String(code || '').trim().toUpperCase();
  return table.values
    .filter(row => String(row[codePos] || '').trim().toUpperCase() === normalizedCode)
    .map(row => row[datePos])
    .filter(value => value instanceof Date && !isNaN(value.getTime()));
}

function FS00P_maxTheoMaAliases_(table, code, aliases) {
  const values = FS00P_chuoiTheoMaAliases_(table, code, aliases);
  return values.length ? Math.max.apply(null, values) : 0;
}

function FS00P_xnpv_(rate, cashFlows, dates) {
  if (!cashFlows.length || cashFlows.length !== dates.length) return 0;
  const firstDate = dates[0];
  return cashFlows.reduce((sum, cashFlow, index) => {
    const years = (dates[index].getTime() - firstDate.getTime()) / 86400000 / 365;
    return sum + FS00_num_(cashFlow) / Math.pow(1 + rate, years);
  }, 0);
}

function FS00P_xirr_(cashFlows, dates) {
  if (!cashFlows.length || cashFlows.length !== dates.length) return 0;
  const hasPositive = cashFlows.some(value => FS00_num_(value) > 0);
  const hasNegative = cashFlows.some(value => FS00_num_(value) < 0);
  if (!hasPositive || !hasNegative) return 0;

  let low = -0.9999;
  let high = 10;
  let lowValue = FS00P_xnpv_(low, cashFlows, dates);
  let highValue = FS00P_xnpv_(high, cashFlows, dates);
  if (!isFinite(lowValue) || !isFinite(highValue) || lowValue * highValue > 0) return 0;

  for (let i = 0; i < 200; i++) {
    const mid = (low + high) / 2;
    const midValue = FS00P_xnpv_(mid, cashFlows, dates);
    if (!isFinite(midValue)) return 0;
    if (Math.abs(midValue) < 0.01) return mid;
    if (lowValue * midValue <= 0) {
      high = mid;
      highValue = midValue;
    } else {
      low = mid;
      lowValue = midValue;
    }
  }
  return (low + high) / 2;
}
