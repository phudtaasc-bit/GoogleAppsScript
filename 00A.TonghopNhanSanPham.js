function FS_lapSheet00_TheoDanhMuc() {
  const result = FS_lapSheet00();
  FS00_capNhatNhanSanPham_();
  FS00_capNhatChiPhiSauVat_();
  FS00_capNhatCoCauNguonVon_();
  FS00_capNhatHoanVonBenVung_();
  return result;
}

function FS00_capNhatNhanSanPham_() {
  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName('01A. Kỹ thuật');
  const summary = ss.getSheetByName('00. Tổng hợp');
  if (!tech || !summary) return;

  const block = FS00_findBlock_(tech, 'SAN_PHAM');
  const values = tech.getRange(block.startRow, 1, block.rowCount, 14).getValues();
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

  const totalRevenueRow = FS00_findSummaryRow_(summary, 'Tổng doanh thu có VAT');
  const totalCostRow = FS00_findSummaryRow_(summary, 'Tổng chi phí có VAT');
  if (!totalRevenueRow || !totalCostRow) return;

  const detailCount = Math.min(products.length, totalCostRow - totalRevenueRow - 1);
  for (let i = 0; i < detailCount; i++) {
    const product = products[i];
    const label = 'Phần ' + product.name + (product.group ? ' - ' + product.group : '');
    summary.getRange(totalRevenueRow + 1 + i, 2).setValue(label);
  }
}

function FS00_capNhatChiPhiSauVat_() {
  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName('01A. Kỹ thuật');
  const costSheet = ss.getSheetByName('03. Chi phí & Vốn');
  const cashSheet = ss.getSheetByName('04. Dòng tiền & Tài trợ');
  const summary = ss.getSheetByName('00. Tổng hợp');

  if (!tech || !costSheet || !cashSheet || !summary) {
    throw new Error('Thiếu sheet nguồn để cập nhật chi phí sau VAT trên Sheet 00.');
  }

  const cost = FS00_readTable_(costSheet);
  const cash = FS00_readTable_(cashSheet);

  FS00_require_(cost.index, [
    'xdtbtruocvat', 'gpmbtruocvat', 'htkttruocvat',
    'tiensddtruocvat', 'tienthuedattruocvat',
    'chiphibanhangtruocvat', 'chiphivanhanhtruocvat',
    'chiphibaotritruocvat', 'chiphiduphongtruocvat',
    'tongchisauvat'
  ], costSheet.getName());
  FS00_require_(cash.index, ['laivay'], cashSheet.getName());

  const vatRates = FS00_docVatRates_(tech);
  const sumCost = key => FS00_sumColumn_(cost, key);
  const interest = FS00_sumColumn_(cash, 'laivay');

  const construction = sumCost('xdtbtruocvat') * (1 + vatRates.construction);
  const clearance = sumCost('gpmbtruocvat') * (1 + vatRates.clearance);
  const landUse = sumCost('tiensddtruocvat') * (1 + vatRates.landUse);
  const landRent = sumCost('tienthuedattruocvat') * (1 + vatRates.landRent);
  const land = landUse + landRent;
  const infrastructure = sumCost('htkttruocvat') * (1 + vatRates.infrastructure);
  const contingency = sumCost('chiphiduphongtruocvat') * (1 + vatRates.contingency);
  const selling = sumCost('chiphibanhangtruocvat') * (1 + vatRates.selling);
  const operating = sumCost('chiphivanhanhtruocvat') * (1 + vatRates.operating);
  const maintenance = sumCost('chiphibaotritruocvat') * (1 + vatRates.maintenance);

  const billion = value => FS00_num_(value) / 1e9;

  summary.getRange('C6').setValue(billion(construction));
  summary.getRange('C7').setValue(billion(clearance));
  summary.getRange('C8').setValue(billion(land));
  summary.getRange('C9').setValue(billion(infrastructure));
  summary.getRange('C10').setValue(billion(contingency));
  summary.getRange('C11').setValue(billion(interest));
  summary.getRange('C12').setFormula('=SUM(C6:C11)');
  summary.getRange('C13').setFormula('=C12-C8');
  summary.getRange('D6:D11').setFormulaR1C1('=IF(R12C3=0;0;RC[-1]/R12C3)');
  summary.getRange('D12').setValue(1);

  const totalCostRow = FS00_findSummaryRow_(summary, 'Tổng chi phí có VAT');
  const coreInvestmentRow = FS00_findSummaryRow_(summary, 'Tổng vốn đầu tư dự án');
  const sellingRow = FS00_findSummaryRow_(summary, 'Chi phí bán hàng');
  const operatingRow = FS00_findSummaryRow_(summary, 'Chi phí vận hành');
  const maintenanceRow = FS00_findSummaryRow_(summary, 'Chi phí bảo trì');

  if (!totalCostRow || !coreInvestmentRow || !sellingRow || !operatingRow || !maintenanceRow) {
    throw new Error('Không xác định đủ các dòng chi phí tại Mục III trên Sheet 00.');
  }

  summary.getRange(coreInvestmentRow, 4).setFormula('=$C$12');
  summary.getRange(sellingRow, 4).setValue(billion(selling));
  summary.getRange(operatingRow, 4).setValue(billion(operating));
  summary.getRange(maintenanceRow, 4).setValue(billion(maintenance));
  summary.getRange(totalCostRow, 4).setFormula(
    '=SUM(D' + coreInvestmentRow + ':D' + maintenanceRow + ')'
  );

  const expectedTotal = FS00_sumColumn_(cost, 'tongchisauvat') + interest;
  const componentTotal = construction + clearance + land + infrastructure + contingency +
    interest + selling + operating + maintenance;

  if (Math.abs(expectedTotal - componentTotal) > 2) {
    throw new Error(
      'Chi phí sau VAT tại Sheet 00 chưa khớp Sheet 03. Sai lệch: ' +
      Math.round(expectedTotal - componentTotal).toLocaleString('vi-VN') + ' đồng.'
    );
  }

  summary.getRange('C6:C13').setNumberFormat('#,##0.0');
  summary.getRange('D6:D12').setNumberFormat('0.0%');
  summary.getRange(totalCostRow, 4, maintenanceRow - totalCostRow + 1, 1)
    .setNumberFormat('#,##0.0');

  SpreadsheetApp.flush();
}

function FS00_capNhatCoCauNguonVon_() {
  const ss = SpreadsheetApp.getActive();
  const summary = ss.getSheetByName('00. Tổng hợp');
  if (!summary) throw new Error('Không tìm thấy Sheet 00. Tổng hợp.');

  SpreadsheetApp.flush();

  const totalInvestment = FS00_num_(summary.getRange('C12').getValue());
  const equity = FS00_num_(summary.getRange('C17').getValue());
  const loan = FS00_num_(summary.getRange('C18').getValue());
  const customerFunding = totalInvestment - equity - loan;

  if (customerFunding < -0.000000001) {
    throw new Error(
      'Vốn CSH và vốn vay vượt Tổng nguồn vốn. Sai lệch: ' +
      Math.round(Math.abs(customerFunding) * 1e9).toLocaleString('vi-VN') + ' đồng.'
    );
  }

  summary.getRange('C19').setValue(Math.max(0, customerFunding));
  summary.getRange('C20').setFormula('=C12');
  summary.getRange('D17').setFormula('=IF(C20=0;0;C17/C20)');
  summary.getRange('D18').setFormula('=IF(C20=0;0;C18/C20)');
  summary.getRange('D19').setFormula('=IF(C20=0;0;C19/C20)');
  summary.getRange('D20').setValue(1);
  summary.getRange('E21').setFormula('=D17*E17+D18*E18+D19*E19');

  summary.getRange('C17:C20').setNumberFormat('#,##0.0');
  summary.getRange('D17:D20').setNumberFormat('0.0%');
  summary.getRange('E21').setNumberFormat('0.00%');

  SpreadsheetApp.flush();
}

function FS00_capNhatHoanVonBenVung_() {
  const ss = SpreadsheetApp.getActive();
  const cashSheet = ss.getSheetByName('04. Dòng tiền & Tài trợ');
  const summary = ss.getSheetByName('00. Tổng hợp');

  if (!cashSheet || !summary) {
    throw new Error('Thiếu Sheet 04 hoặc Sheet 00 để tính thời gian hoàn vốn.');
  }

  const cash = FS00_readTable_(cashSheet);
  FS00_require_(cash.index, ['fcff', 'fcfe'], cashSheet.getName());

  const fcff = FS00_columnValues_(cash, 'fcff');
  const fcfe = FS00_columnValues_(cash, 'fcfe');

  const projectRow = FS00_findSummaryRow_(summary, 'Thời gian hoàn vốn dự án');
  const equityRow = FS00_findSummaryRowByAliases_(summary, [
    'Thời gian hoàn vốn - Vốn CSH',
    'Thời gian hoàn vốn vốn CSH'
  ]);

  if (!projectRow || !equityRow) {
    throw new Error('Không xác định được các dòng thời gian hoàn vốn trên Sheet 00.');
  }

  summary.getRange(projectRow, 4)
    .setValue(FS00_tinhHoanVonBenVung_(fcff))
    .setNumberFormat('0.00');

  summary.getRange(equityRow, 4)
    .setValue(FS00_tinhHoanVonBenVung_(fcfe))
    .setNumberFormat('0.00');
}

function FS00_tinhHoanVonBenVung_(cashFlows) {
  const tolerance = 1;
  const cumulative = [];
  let runningTotal = 0;
  let lastNegativeIndex = -1;

  cashFlows.forEach((value, index) => {
    runningTotal += FS00_num_(value);
    cumulative.push(runningTotal);
    if (runningTotal < -tolerance) lastNegativeIndex = index;
  });

  if (!cashFlows.length) return 0;
  if (lastNegativeIndex < 0) return 0;

  const recoveryIndex = lastNegativeIndex + 1;
  if (recoveryIndex >= cashFlows.length) return 0;

  for (let index = recoveryIndex; index < cumulative.length; index++) {
    if (cumulative[index] < -tolerance) return 0;
  }

  const previousCumulative = cumulative[lastNegativeIndex];
  const recoveryCashFlow = FS00_num_(cashFlows[recoveryIndex]);

  if (recoveryCashFlow <= 0) return recoveryIndex;

  const fraction = Math.min(1, Math.max(0, Math.abs(previousCumulative) / recoveryCashFlow));
  return lastNegativeIndex + 1 + fraction;
}

function FS00_docVatRates_(sheet) {
  const block = FS00_findBlock_(sheet, 'CHI_PHI_CHUNG');
  const values = sheet.getRange(block.startRow, 1, block.rowCount, 6).getValues();
  const map = {};

  values.forEach(row => {
    const name = String(row[0] || '').trim();
    if (!name) return;
    map[FS00_key_(name)] = FS00_docRate_(row[2]);
  });

  const get = (aliases, fallback) => {
    for (const alias of aliases) {
      const key = FS00_key_(alias);
      if (map[key] != null) return map[key];
    }
    return fallback;
  };

  const rentRates = [
    get(['Tiền thuê đất Chung cư'], 0),
    get(['Tiền thuê đất TMDV'], 0),
    get(['Tiền thuê đất Chợ'], 0)
  ];
  const nonZeroRentRates = rentRates.filter(rate => rate !== 0);

  return {
    construction: get(['Chi phí XD/TB/khác', 'Chi phí XD/TB'], 0.08),
    clearance: get(['Chi phí GPMB'], 0),
    landUse: get(['Tiền SDĐ Liền kề', 'Tiền SDD Liền kề'], 0),
    landRent: nonZeroRentRates.length ? nonZeroRentRates[0] : 0,
    infrastructure: get(['Chi phí HTKT'], 0.08),
    contingency: get(['Chi phí dự phòng'], 0.08),
    selling: get(['Chi phí bán hàng'], 0.08),
    operating: get(['Chi phí vận hành'], 0.08),
    maintenance: get(['Chi phí bảo trì'], 0.08)
  };
}

function FS00_docRate_(value) {
  const number = FS00_num_(value);
  return Math.abs(number) > 1 ? number / 100 : number;
}
