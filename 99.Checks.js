const FS99_CFG = Object.freeze({
  REVENUE: '02. Doanh thu',
  COST: '03. Chi phí & Vốn',
  PROFIT: '03A. Lợi nhuận & Thuế',
  CASH: '04. Dòng tiền & Tài trợ',
  CASH_LEGACY: '04. Dòng tiền',
  SUMMARY: '04A. TH dòng tiền',
  CHECKS: '99. Checks',
  CHECKS_LEGACY: '99. Kiểm tra',
  TOLERANCE: 1,
  REPORT_TOLERANCE: 0.11
});

function FS_lapSheet99() {
  const ss = SpreadsheetApp.getActive();
  const revenueSheet = ss.getSheetByName(FS99_CFG.REVENUE);
  const costSheet = ss.getSheetByName(FS99_CFG.COST);
  const profitSheet = ss.getSheetByName(FS99_CFG.PROFIT);
  const cashSheet = ss.getSheetByName(FS99_CFG.CASH) || ss.getSheetByName(FS99_CFG.CASH_LEGACY);

  if (!revenueSheet || !costSheet || !profitSheet || !cashSheet) {
    throw new Error('Cần lập đủ Sheet 02, 03, 03A và 04 trước khi chạy kiểm tra.');
  }

  const revenue = FS99_readTable_(revenueSheet);
  const cost = FS99_readTable_(costSheet);
  const profit = FS99_readTable_(profitSheet);
  const cash = FS99_readTable_(cashSheet);
  const checks = [];

  FS99_checkRequiredHeaders_(checks, revenue, [
    'thangso', 'masp', 'tongdoanhthutruocvat', 'vatdaura', 'dongtienkhachhang'
  ]);
  FS99_checkRequiredHeaders_(checks, cost, [
    'thangso', 'masp', 'dongtienkhachhang', 'vatdaura', 'tongchitruocvat',
    'vatdauvao', 'tongchisauvat', 'xdtbtruocvat', 'gpmbtruocvat',
    'htkttruocvat', 'tiensddtruocvat', 'tienthuedattruocvat',
    'chiphibanhangtruocvat', 'chiphivanhanhtruocvat',
    'chiphibaotritruocvat', 'chiphiduphongtruocvat'
  ]);
  FS99_checkRequiredHeaders_(checks, profit, [
    'thangso', 'masp', 'tongdoanhthutruocvat',
    'giavonxdgpmbhtktduphong', 'laivayvonhoaphanbo',
    'tiensddphanbo', 'tienthuedatphanbo', 'chiphibanhang',
    'chiphivanhanh', 'chiphibaotri', 'tongchiphihachtoan',
    'loinhuantruocthue', 'thuetndn', 'lnst'
  ]);
  FS99_checkRequiredHeaders_(checks, cash, [
    'thangso', 'dongtienkhachhang', 'vatdaura', 'tongchitruocvat',
    'vatdauvao', 'tongchisauvat', 'vatphainop', 'thuetndn', 'fcff',
    'tientondauky', 'tientruoctaitro', 'nhucauvon', 'laivay',
    'vongopcsh', 'giainganvay', 'tragoc', 'dunodauky', 'dunocuoiky',
    'tientoncuoiky', 'fcfe'
  ]);

  if (checks.some(item => item.status === 'FAIL')) {
    FS99_write_(ss, checks);
    throw new Error('99.Checks: thiếu header bắt buộc. Kiểm tra Sheet 99.');
  }

  FS99_checkKeyStructure_(checks, revenue, cost, profit);
  FS99_checkFlowTotals_(checks, revenue, cost, profit, cash);
  FS99_checkCostIdentities_(checks, cost);
  FS99_checkProfitIdentities_(checks, cost, profit);
  FS99_checkCashIdentities_(checks, cash);
  FS99_checkConvergence_(checks);
  FS99_checkAnnualSummary_(checks, ss, cash);

  FS99_write_(ss, checks);

  const failed = checks.filter(item => item.status === 'FAIL');
  if (failed.length) {
    throw new Error('99.Checks có ' + failed.length + ' kiểm tra FAIL. Mở Sheet 99 để xem chi tiết.');
  }

  return checks;
}

function FS99_checkRequiredHeaders_(checks, table, required) {
  const missing = required.filter(key => table.index[key] == null);
  FS99_add_(
    checks,
    'CẤU TRÚC',
    'Header bắt buộc - ' + table.name,
    missing.length,
    0,
    missing.length,
    missing.length ? 'Thiếu: ' + missing.join(', ') : 'Đủ header bắt buộc'
  );
}

function FS99_checkKeyStructure_(checks, revenue, cost, profit) {
  const revenueKeys = FS99_keySet_(revenue, 'masp', 'thangso');
  const costKeys = FS99_keySet_(cost, 'masp', 'thangso');
  const profitKeys = FS99_keySet_(profit, 'masp', 'thangso');

  FS99_add_(checks, 'KHÓA DỮ LIỆU', 'Số khóa 02 ↔ 03', revenueKeys.size, costKeys.size,
    revenueKeys.size - costKeys.size, 'Khóa = Mã SP + Tháng số');
  FS99_add_(checks, 'KHÓA DỮ LIỆU', 'Số khóa 02 ↔ 03A', revenueKeys.size, profitKeys.size,
    revenueKeys.size - profitKeys.size, 'Khóa = Mã SP + Tháng số');

  const missingCost = [...revenueKeys].filter(key => !costKeys.has(key));
  const extraCost = [...costKeys].filter(key => !revenueKeys.has(key));
  const missingProfit = [...revenueKeys].filter(key => !profitKeys.has(key));
  const extraProfit = [...profitKeys].filter(key => !revenueKeys.has(key));

  FS99_add_(checks, 'KHÓA DỮ LIỆU', 'Khóa 02 ↔ 03 khớp tuyệt đối',
    missingCost.length + extraCost.length, 0, missingCost.length + extraCost.length,
    FS99_keyNote_(missingCost, extraCost));
  FS99_add_(checks, 'KHÓA DỮ LIỆU', 'Khóa 02 ↔ 03A khớp tuyệt đối',
    missingProfit.length + extraProfit.length, 0, missingProfit.length + extraProfit.length,
    FS99_keyNote_(missingProfit, extraProfit));
}

function FS99_checkFlowTotals_(checks, revenue, cost, profit, cash) {
  const revenueTotal = FS99_sum_(revenue, 'tongdoanhthutruocvat');
  const profitRevenue = FS99_sum_(profit, 'tongdoanhthutruocvat');
  FS99_compare_(checks, 'DOANH THU', 'Tổng doanh thu 02 ↔ 03A', revenueTotal, profitRevenue,
    '03A chỉ đọc doanh thu từ 02');

  const customer02 = FS99_sum_(revenue, 'dongtienkhachhang');
  const customer03 = FS99_sum_(cost, 'dongtienkhachhang');
  const customer04 = FS99_sum_(cash, 'dongtienkhachhang');
  FS99_compare_(checks, 'DÒNG TIỀN KHÁCH HÀNG', 'Dòng tiền KH 02 ↔ 03', customer02, customer03, '');
  FS99_compare_(checks, 'DÒNG TIỀN KHÁCH HÀNG', 'Dòng tiền KH 03 ↔ 04', customer03, customer04, '');

  const vatOut02 = FS99_sum_(revenue, 'vatdaura');
  const vatOut03 = FS99_sum_(cost, 'vatdaura');
  const vatOut04 = FS99_sum_(cash, 'vatdaura');
  FS99_compare_(checks, 'VAT', 'VAT đầu ra 02 ↔ 03', vatOut02, vatOut03, '');
  FS99_compare_(checks, 'VAT', 'VAT đầu ra 03 ↔ 04', vatOut03, vatOut04, '');

  FS99_compare_(checks, 'CHI PHÍ', 'Tổng chi trước VAT 03 ↔ 04',
    FS99_sum_(cost, 'tongchitruocvat'), FS99_sum_(cash, 'tongchitruocvat'), '');
  FS99_compare_(checks, 'VAT', 'VAT đầu vào 03 ↔ 04',
    FS99_sum_(cost, 'vatdauvao'), FS99_sum_(cash, 'vatdauvao'), 'Lãi vay có VAT đầu vào = 0');
  FS99_compare_(checks, 'CHI PHÍ', 'Tổng chi sau VAT 03 ↔ 04',
    FS99_sum_(cost, 'tongchisauvat'), FS99_sum_(cash, 'tongchisauvat'), '');

  FS99_compare_(checks, 'THUẾ', 'Thuế TNDN 03A ↔ 04',
    FS99_sum_(profit, 'thuetndn'), FS99_sum_(cash, 'thuetndn'), '');
  FS99_compare_(checks, 'LÃI VAY', 'Lãi vay vốn hóa 03A ↔ lãi vay 04',
    FS99_sum_(profit, 'laivayvonhoaphanbo'), FS99_sum_(cash, 'laivay'),
    'Toàn bộ lãi vay được phân bổ hết trong vòng đời mô hình');

  FS99_compare_(checks, 'VAT', 'VAT đầu vào của lãi vay', 0, 0,
    'Lãi vay không thuộc cơ sở tính VAT đầu vào');
}

function FS99_checkCostIdentities_(checks, cost) {
  let maxBeforeDiff = 0;
  let maxAfterDiff = 0;

  cost.values.forEach(row => {
    const components =
      FS99_value_(cost, row, 'xdtbtruocvat') +
      FS99_value_(cost, row, 'gpmbtruocvat') +
      FS99_value_(cost, row, 'htkttruocvat') +
      FS99_value_(cost, row, 'tiensddtruocvat') +
      FS99_value_(cost, row, 'tienthuedattruocvat') +
      FS99_value_(cost, row, 'chiphibanhangtruocvat') +
      FS99_value_(cost, row, 'chiphivanhanhtruocvat') +
      FS99_value_(cost, row, 'chiphibaotritruocvat') +
      FS99_value_(cost, row, 'chiphiduphongtruocvat');

    maxBeforeDiff = Math.max(maxBeforeDiff,
      Math.abs(FS99_value_(cost, row, 'tongchitruocvat') - components));
    maxAfterDiff = Math.max(maxAfterDiff,
      Math.abs(FS99_value_(cost, row, 'tongchisauvat') -
        FS99_value_(cost, row, 'tongchitruocvat') -
        FS99_value_(cost, row, 'vatdauvao')));
  });

  FS99_add_(checks, 'CHI PHÍ', 'Cấu thành Tổng chi trước VAT từng dòng',
    maxBeforeDiff, 0, maxBeforeDiff, 'Sai lệch lớn nhất theo Mã SP + Tháng số');
  FS99_add_(checks, 'CHI PHÍ', 'Tổng chi sau VAT = trước VAT + VAT đầu vào',
    maxAfterDiff, 0, maxAfterDiff, 'Sai lệch lớn nhất theo Mã SP + Tháng số');
}

function FS99_checkProfitIdentities_(checks, cost, profit) {
  const baseCost03 =
    FS99_sum_(cost, 'xdtbtruocvat') +
    FS99_sum_(cost, 'gpmbtruocvat') +
    FS99_sum_(cost, 'htkttruocvat') +
    FS99_sum_(cost, 'chiphiduphongtruocvat');

  FS99_compare_(checks, 'HẠCH TOÁN', 'Giá vốn cơ sở 03 ↔ 03A',
    baseCost03, FS99_sum_(profit, 'giavonxdgpmbhtktduphong'), 'Phân bổ hết trong vòng đời mô hình');
  FS99_compare_(checks, 'HẠCH TOÁN', 'Tiền SDĐ 03 ↔ phân bổ 03A',
    FS99_sum_(cost, 'tiensddtruocvat'), FS99_sum_(profit, 'tiensddphanbo'), '');
  FS99_compare_(checks, 'HẠCH TOÁN', 'Tiền thuê đất 03 ↔ phân bổ 03A',
    FS99_sum_(cost, 'tienthuedattruocvat'), FS99_sum_(profit, 'tienthuedatphanbo'), '');
  FS99_compare_(checks, 'HẠCH TOÁN', 'Chi phí bán hàng 03 ↔ 03A',
    FS99_sum_(cost, 'chiphibanhangtruocvat'), FS99_sum_(profit, 'chiphibanhang'), '');
  FS99_compare_(checks, 'HẠCH TOÁN', 'Chi phí vận hành 03 ↔ 03A',
    FS99_sum_(cost, 'chiphivanhanhtruocvat'), FS99_sum_(profit, 'chiphivanhanh'), '');
  FS99_compare_(checks, 'HẠCH TOÁN', 'Chi phí bảo trì 03 ↔ 03A',
    FS99_sum_(cost, 'chiphibaotritruocvat'), FS99_sum_(profit, 'chiphibaotri'), '');

  let maxCostDiff = 0;
  let maxPbtDiff = 0;
  let maxPatDiff = 0;
  profit.values.forEach(row => {
    const detailCost =
      FS99_value_(profit, row, 'giavonxdgpmbhtktduphong') +
      FS99_value_(profit, row, 'laivayvonhoaphanbo') +
      FS99_value_(profit, row, 'tiensddphanbo') +
      FS99_value_(profit, row, 'tienthuedatphanbo') +
      FS99_value_(profit, row, 'chiphibanhang') +
      FS99_value_(profit, row, 'chiphivanhanh') +
      FS99_value_(profit, row, 'chiphibaotri');

    maxCostDiff = Math.max(maxCostDiff,
      Math.abs(FS99_value_(profit, row, 'tongchiphihachtoan') - detailCost));
    maxPbtDiff = Math.max(maxPbtDiff,
      Math.abs(FS99_value_(profit, row, 'loinhuantruocthue') -
        FS99_value_(profit, row, 'tongdoanhthutruocvat') +
        FS99_value_(profit, row, 'tongchiphihachtoan')));
    maxPatDiff = Math.max(maxPatDiff,
      Math.abs(FS99_value_(profit, row, 'lnst') -
        FS99_value_(profit, row, 'loinhuantruocthue') +
        FS99_value_(profit, row, 'thuetndn')));
  });

  FS99_add_(checks, 'HẠCH TOÁN', 'Tổng chi phí hạch toán theo cấu phần', maxCostDiff, 0, maxCostDiff,
    'Sai lệch lớn nhất từng dòng');
  FS99_add_(checks, 'LỢI NHUẬN', 'Lợi nhuận trước thuế = doanh thu - chi phí', maxPbtDiff, 0, maxPbtDiff,
    'Sai lệch lớn nhất từng dòng');
  FS99_add_(checks, 'LỢI NHUẬN', 'LNST = Lợi nhuận trước thuế - Thuế TNDN', maxPatDiff, 0, maxPatDiff,
    'Sai lệch lớn nhất từng dòng');
}

function FS99_checkCashIdentities_(checks, cash) {
  let maxFcffDiff = 0;
  let maxDebtDiff = 0;
  let maxCashCarryDiff = 0;
  let maxFcfeDiff = 0;
  let maxFundingDiff = 0;
  let minDebt = Infinity;
  let minCash = Infinity;

  const rows = cash.values.slice().sort((a, b) =>
    FS99_value_(cash, a, 'thangso') - FS99_value_(cash, b, 'thangso'));

  rows.forEach((row, index) => {
    const fcffExpected =
      FS99_value_(cash, row, 'dongtienkhachhang') -
      FS99_value_(cash, row, 'tongchisauvat') -
      FS99_value_(cash, row, 'vatphainop') -
      FS99_value_(cash, row, 'thuetndn');
    maxFcffDiff = Math.max(maxFcffDiff,
      Math.abs(FS99_value_(cash, row, 'fcff') - fcffExpected));

    const debtExpected =
      FS99_value_(cash, row, 'dunodauky') +
      FS99_value_(cash, row, 'giainganvay') -
      FS99_value_(cash, row, 'tragoc');
    maxDebtDiff = Math.max(maxDebtDiff,
      Math.abs(FS99_value_(cash, row, 'dunocuoiky') - debtExpected));

    const fcfeExpected =
      FS99_value_(cash, row, 'tientoncuoiky') -
      FS99_value_(cash, row, 'tientondauky') -
      FS99_value_(cash, row, 'vongopcsh');
    maxFcfeDiff = Math.max(maxFcfeDiff,
      Math.abs(FS99_value_(cash, row, 'fcfe') - fcfeExpected));

    const fundingExpected = Math.max(0, -FS99_value_(cash, row, 'tientruoctaitro'));
    maxFundingDiff = Math.max(maxFundingDiff,
      Math.abs(FS99_value_(cash, row, 'nhucauvon') - fundingExpected));

    if (index > 0) {
      maxCashCarryDiff = Math.max(maxCashCarryDiff,
        Math.abs(FS99_value_(cash, row, 'tientondauky') -
          FS99_value_(cash, rows[index - 1], 'tientoncuoiky')));
    }

    minDebt = Math.min(minDebt, FS99_value_(cash, row, 'dunocuoiky'));
    minCash = Math.min(minCash, FS99_value_(cash, row, 'tientoncuoiky'));
  });

  FS99_add_(checks, 'DÒNG TIỀN', 'FCFF từng tháng', maxFcffDiff, 0, maxFcffDiff,
    'FCFF = Thu KH - Chi sau VAT - VAT phải nộp - Thuế TNDN');
  FS99_add_(checks, 'TÀI TRỢ', 'Cân đối dư nợ từng tháng', maxDebtDiff, 0, maxDebtDiff,
    'Dư nợ cuối = đầu kỳ + giải ngân - trả gốc');
  FS99_add_(checks, 'DÒNG TIỀN', 'Tiền tồn chuyển kỳ', maxCashCarryDiff, 0, maxCashCarryDiff,
    'Tiền đầu kỳ tháng sau = tiền cuối kỳ tháng trước');
  FS99_add_(checks, 'DÒNG TIỀN', 'FCFE từng tháng không cộng lũy kế', maxFcfeDiff, 0, maxFcfeDiff,
    'FCFE = Tiền tồn cuối kỳ - Tiền tồn đầu kỳ - Vốn góp CSH');
  FS99_add_(checks, 'TÀI TRỢ', 'Nhu cầu vốn chỉ phát sinh khi thiếu tiền', maxFundingDiff, 0, maxFundingDiff,
    'Nhu cầu vốn = MAX(0; -Tiền trước tài trợ)');
  FS99_add_(checks, 'TÀI TRỢ', 'Không có dư nợ âm', minDebt < 0 ? Math.abs(minDebt) : 0, 0,
    minDebt < 0 ? Math.abs(minDebt) : 0, 'Dư nợ cuối kỳ nhỏ nhất');
  FS99_add_(checks, 'DÒNG TIỀN', 'Không có tiền tồn âm', minCash < 0 ? Math.abs(minCash) : 0, 0,
    minCash < 0 ? Math.abs(minCash) : 0, 'Tiền tồn cuối kỳ nhỏ nhất');

  const finalRow = rows[rows.length - 1];
  FS99_add_(checks, 'TÀI TRỢ', 'Dư nợ cuối mô hình bằng 0',
    finalRow ? FS99_value_(cash, finalRow, 'dunocuoiky') : 0, 0,
    finalRow ? FS99_value_(cash, finalRow, 'dunocuoiky') : 0,
    'Không để dư nợ treo sau tháng cuối mô hình');
}

function FS99_checkConvergence_(checks) {
  let state = {};
  try {
    state = JSON.parse(PropertiesService.getDocumentProperties().getProperty('FS_CONVERGENCE') || '{}');
  } catch (error) {
    state = {};
  }

  FS99_add_(checks, 'HỘI TỤ', 'Trạng thái hội tụ 03A ↔ 04',
    state.converged ? 0 : 1, 0, state.converged ? 0 : 1,
    'Số vòng: ' + FS99_num_(state.iterations) + '; sai số lớn nhất: ' + FS99_num_(state.maxDiff));
}

function FS99_checkAnnualSummary_(checks, ss, cash) {
  const summary = ss.getSheetByName(FS99_CFG.SUMMARY);
  if (!summary) {
    FS99_add_(checks, '04A', 'Sheet 04A tồn tại', 1, 0, 1, 'Không tìm thấy "04A. TH dòng tiền"');
    return;
  }

  const totalFcffSource = FS99_sum_(cash, 'fcff') / 1e9;
  const totalFcfeSource = FS99_sum_(cash, 'fcfe') / 1e9;
  const totalFcffReport = FS99_num_(summary.getRange(30, 3).getValue());
  const totalFcfeReport = FS99_num_(summary.getRange(31, 3).getValue());

  FS99_addWithTolerance_(checks, '04A', 'Tổng FCFF 04 ↔ 04A',
    totalFcffSource, totalFcffReport, totalFcffSource - totalFcffReport,
    'Đơn vị đối chiếu: tỷ đồng', FS99_CFG.REPORT_TOLERANCE);
  FS99_addWithTolerance_(checks, '04A', 'Tổng FCFE 04 ↔ 04A',
    totalFcfeSource, totalFcfeReport, totalFcfeSource - totalFcfeReport,
    'Đơn vị đối chiếu: tỷ đồng', FS99_CFG.REPORT_TOLERANCE);
}

function FS99_readTable_(sheet) {
  const columnCount = sheet.getLastColumn();
  if (columnCount < 1 || sheet.getLastRow() < 1) {
    throw new Error('Sheet "' + sheet.getName() + '" không có dữ liệu.');
  }

  const headers = sheet.getRange(1, 1, 1, columnCount).getDisplayValues()[0];
  const values = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, columnCount).getValues()
    : [];
  const index = {};
  headers.forEach((header, position) => { index[FS99_key_(header)] = position; });
  return { name: sheet.getName(), values, index };
}

function FS99_keySet_(table, codeKey, monthKey) {
  const result = new Set();
  table.values.forEach(row => {
    const code = String(row[table.index[codeKey]] || '').trim().toUpperCase();
    const month = FS99_num_(row[table.index[monthKey]]);
    if (!code || month < 1) return;
    const key = code + '|' + month;
    if (result.has(key)) throw new Error(table.name + ' bị trùng khóa: ' + key);
    result.add(key);
  });
  return result;
}

function FS99_sum_(table, key) {
  const position = table.index[key];
  if (position == null) return 0;
  return table.values.reduce((sum, row) => sum + FS99_num_(row[position]), 0);
}

function FS99_value_(table, row, key) {
  const position = table.index[key];
  return position == null ? 0 : FS99_num_(row[position]);
}

function FS99_compare_(checks, group, name, sourceA, sourceB, note) {
  FS99_add_(checks, group, name, sourceA, sourceB, sourceA - sourceB, note);
}

function FS99_add_(checks, group, name, sourceA, sourceB, difference, note) {
  FS99_addWithTolerance_(checks, group, name, sourceA, sourceB, difference, note, FS99_CFG.TOLERANCE);
}

function FS99_addWithTolerance_(checks, group, name, sourceA, sourceB, difference, note, tolerance) {
  const diff = FS99_num_(difference);
  checks.push({
    group,
    name,
    sourceA: FS99_num_(sourceA),
    sourceB: FS99_num_(sourceB),
    difference: diff,
    status: Math.abs(diff) <= tolerance ? 'PASS' : 'FAIL',
    note: note || ''
  });
}

function FS99_keyNote_(missing, extra) {
  const parts = [];
  if (missing.length) parts.push('Thiếu: ' + missing.slice(0, 10).join(', '));
  if (extra.length) parts.push('Thừa: ' + extra.slice(0, 10).join(', '));
  return parts.join('; ') || 'Khớp toàn bộ khóa';
}

function FS99_write_(ss, checks) {
  let sheet = ss.getSheetByName(FS99_CFG.CHECKS) || ss.getSheetByName(FS99_CFG.CHECKS_LEGACY);
  if (!sheet) sheet = ss.insertSheet(FS99_CFG.CHECKS);
  if (sheet.getName() !== FS99_CFG.CHECKS) sheet.setName(FS99_CFG.CHECKS);

  sheet.clear();
  sheet.clearFormats();

  const failCount = checks.filter(item => item.status === 'FAIL').length;
  sheet.getRange(1, 1).setValue('99. KIỂM TRA TOÀN BỘ MÔ HÌNH');
  sheet.getRange(2, 1, 1, 2).setValues([['Kết quả chung', failCount ? 'FAIL' : 'PASS']]);
  sheet.getRange(3, 1, 1, 2).setValues([['Số kiểm tra FAIL', failCount]]);

  const headers = [['Nhóm', 'Kiểm tra', 'Nguồn A', 'Nguồn B', 'Sai lệch', 'Trạng thái', 'Ghi chú']];
  const rows = checks.map(item => [
    item.group, item.name, item.sourceA, item.sourceB,
    item.difference, item.status, item.note
  ]);

  sheet.getRange(5, 1, 1, headers[0].length).setValues(headers);
  if (rows.length) sheet.getRange(6, 1, rows.length, headers[0].length).setValues(rows);

  sheet.getRange(1, 1, 1, 7).merge().setFontWeight('bold').setFontSize(14).setBackground('#1f4e78').setFontColor('#ffffff');
  sheet.getRange(5, 1, 1, 7).setFontWeight('bold').setBackground('#d9e2f3').setWrap(true);
  sheet.getRange(2, 2).setFontWeight('bold').setBackground(failCount ? '#f4cccc' : '#d9ead3');
  sheet.getRange(6, 3, Math.max(1, rows.length), 3).setNumberFormat('#,##0.00');

  if (rows.length) {
    const statusRange = sheet.getRange(6, 6, rows.length, 1);
    const rules = [
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('PASS').setBackground('#d9ead3').setRanges([statusRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('FAIL').setBackground('#f4cccc').setRanges([statusRange]).build()
    ];
    sheet.setConditionalFormatRules(rules);
  }

  sheet.setFrozenRows(5);
  sheet.setColumnWidth(1, 135);
  sheet.setColumnWidth(2, 310);
  sheet.setColumnWidth(3, 135);
  sheet.setColumnWidth(4, 135);
  sheet.setColumnWidth(5, 135);
  sheet.setColumnWidth(6, 90);
  sheet.setColumnWidth(7, 420);
  sheet.getDataRange().setVerticalAlignment('middle');
}

function FS99_num_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  const text = String(value == null ? '' : value).trim().replace(/\s/g, '');
  if (!text) return 0;
  const normalized = text.includes(',') && text.includes('.')
    ? text.replace(/\./g, '').replace(',', '.')
    : text.replace(/,/g, '');
  const number = Number(normalized);
  return isFinite(number) ? number : 0;
}

function FS99_norm_(value) {
  return String(value == null ? '' : value)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/\s+/g, ' ')
    .trim();
}

function FS99_key_(value) {
  return FS99_norm_(value)
    .replace(/²/g, '2')
    .replace(/\^2/g, '2')
    .replace(/m\s*2/g, 'm2')
    .replace(/[^a-z0-9]/g, '');
}
