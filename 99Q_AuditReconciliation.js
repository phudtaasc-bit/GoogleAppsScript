/*************************************************
 * 99Q_AuditReconciliation.js
 * Audit phân tầng, xác định chính xác sheet/module gây sai lệch.
 *
 * Tầng 1: nguồn chi tiết Sheet 02 -> 03 -> 04.
 * Tầng 2: Sheet 04A đối chiếu trực tiếp với Sheet 04.
 * Tầng 3: Sheet 00 đối chiếu theo tên chỉ tiêu động.
 * Tầng 4: phương trình mục 14.
 * Tầng 5: phân rã chính xác Tổng chi phí có VAT Sheet 00.
 *************************************************/

function FS99P_buoc5_KiemTraNhanh() {
  return FS99Q_chayAuditReconciliation();
}

function FS99Q_chayAuditReconciliation() {
  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName('01. Kỹ thuật');
  const sh02 = ss.getSheetByName('02. Doanh thu');
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');
  const sh04A = ss.getSheetByName('04A. TH dòng tiền');
  const sh00 = ss.getSheetByName('00. Tổng hợp');
  let shC = ss.getSheetByName('99. Checks');

  if (!tech || !sh02 || !sh03 || !sh04 || !sh00) {
    throw new Error('Thiếu Sheet 00, 01. Kỹ thuật, 02, 03 hoặc 04 để audit.');
  }
  if (!shC) shC = ss.insertSheet('99. Checks');

  SpreadsheetApp.flush();

  const TOL_DONG = 1;
  const TOL_TY = 0.001;
  const results = [];

  const add = (layer, test, a, b, note, fix, tol, forceStatus) => {
    const av = FS99Q_num_(a);
    const bv = FS99Q_num_(b);
    const diff = av - bv;
    const limit = tol == null ? TOL_DONG : tol;
    const status = forceStatus || (Math.abs(diff) <= limit ? 'PASS' : 'FAIL');
    results.push([
      layer, test, av, bv, diff, status, note || '',
      status === 'FAIL' ? (fix || 'Kiểm tra nguồn hoặc công thức của chỉ tiêu này.') : ''
    ]);
  };

  const info = (layer, test, value, note) => {
    results.push([layer, test, FS99Q_num_(value), FS99Q_num_(value), 0, 'INFO', note || '', '']);
  };

  const meta02 = FS99Q_detectHeader_(sh02, ['Tháng số', 'Dòng tiền huy động từ KH', 'Thuế TNDN tạm tính']);
  const meta03 = FS99Q_detectHeader_(sh03, ['Tháng số', 'Tổng chi sau VAT', 'Thuế TNDN']);
  const meta04 = FS99Q_detectHeader_(sh04, ['Tháng số', 'FCFF_TIPV', 'Thuế TNDN']);

  const sum02 = name => FS99Q_sumNamed_(sh02, meta02, name);
  const sum03 = name => FS99Q_sumNamed_(sh03, meta03, name);
  const sum04 = name => FS99Q_sumNamed_(sh04, meta04, name);

  info('TẦNG 1 - CẤU TRÚC', 'Dòng tiêu đề Sheet 02', meta02.headerRow,
    'Phát hiện tự động; dữ liệu bắt đầu tại dòng ' + (meta02.headerRow + 1) + '.');
  info('TẦNG 1 - CẤU TRÚC', 'Dòng tiêu đề Sheet 03', meta03.headerRow,
    'Phát hiện tự động; dữ liệu bắt đầu tại dòng ' + (meta03.headerRow + 1) + '.');
  info('TẦNG 1 - CẤU TRÚC', 'Dòng tiêu đề Sheet 04', meta04.headerRow,
    'Phát hiện tự động; dữ liệu bắt đầu tại dòng ' + (meta04.headerRow + 1) + '.');

  // TẦNG 1 - NGUỒN CHI TIẾT
  add('TẦNG 1 - DOANH THU', 'Sheet 02: Tổng doanh thu trước VAT = bán + thuê',
    sum02('Tổng doanh thu trước VAT'),
    sum02('Doanh thu bán trước VAT') + sum02('Doanh thu thuê trước VAT'),
    'Kiểm tra nội tại Sheet 02.', 'Sửa hàm lập Sheet 02 tại phần tổng doanh thu.', TOL_DONG);

  add('TẦNG 1 - DOANH THU', 'Dòng tiền khách hàng Sheet 02 = Sheet 03',
    sum02('Dòng tiền huy động từ KH'), sum03('Dòng tiền huy động từ KH'),
    'Kiểm tra luồng dữ liệu 02 -> 03.',
    'Sửa hàm cập nhật nguồn Sheet 03 hoặc cột Dòng tiền huy động từ KH.', TOL_DONG);

  const costNames = [
    'Chi XD/TB/khác trước VAT', 'Chi GPMB trước VAT', 'Tiền SDĐ/thuê đất trước VAT',
    'Chi HTKT trước VAT', 'Chi phí bán hàng trước VAT', 'Chi phí dự phòng trước VAT',
    'Chi phí vận hành thuê trước VAT', 'Chi phí bảo trì trước VAT'
  ];
  const parts03 = costNames.reduce((s, n) => s + sum03(n), 0);

  add('TẦNG 1 - CHI PHÍ', 'Sheet 03: Tổng chi trước VAT = tổng 8 cấu phần',
    sum03('Tổng chi trước VAT'), parts03,
    'XD/TB + GPMB + đất + HTKT + bán hàng + dự phòng + vận hành + bảo trì.',
    'Sửa hàm lập/cập nhật Sheet 03; kiểm tra cột Tổng chi trước VAT.', TOL_DONG);

  add('TẦNG 1 - CHI PHÍ', 'Sheet 03: Tổng chi sau VAT = trước VAT + VAT đầu vào',
    sum03('Tổng chi sau VAT'), sum03('Tổng chi trước VAT') + sum03('VAT đầu vào'),
    'Kiểm tra nội tại Sheet 03.',
    'Sửa cột Tổng chi sau VAT hoặc VAT đầu vào tại Sheet 03.', TOL_DONG);

  add('TẦNG 1 - THUẾ', 'Thuế TNDN Sheet 02 = Sheet 03',
    sum02('Thuế TNDN tạm tính'), sum03('Thuế TNDN'),
    'Kiểm tra luồng thuế 02 -> 03.',
    'Sửa module đồng bộ Thuế TNDN từ Sheet 02 sang Sheet 03.', TOL_DONG);

  add('TẦNG 1 - THUẾ', 'Thuế TNDN Sheet 03 = Sheet 04',
    sum03('Thuế TNDN'), sum04('Thuế TNDN'),
    'Kiểm tra luồng thuế 03 -> 04.',
    'Sửa module lập dòng tiền hoặc đồng bộ Thuế TNDN sang Sheet 04.', TOL_DONG);

  add('TẦNG 1 - DÒNG TIỀN', 'FCFF Sheet 04 = Dòng tiền trước tài trợ',
    sum04('FCFF_TIPV'), sum04('Dòng tiền trước tài trợ'),
    'FCFF giữ nguyên logic đã chốt.', 'Sửa công thức FCFF_TIPV tại Sheet 04.', TOL_DONG);

  add('TẦNG 1 - DÒNG TIỀN', 'FCFE = FCFF + giải ngân vay - trả gốc',
    sum04('FCFE khả dụng cho CSH'),
    sum04('FCFF_TIPV') + sum04('Giải ngân vay') - sum04('Trả gốc'),
    'Quan hệ FCFE đã chốt.', 'Sửa công thức FCFE khả dụng cho CSH tại Sheet 04.', TOL_DONG);

  // TẦNG 2 - SHEET 04A
  if (sh04A) {
    const annual = FS99Q_readAnnualTotals_(sh04A);
    FS99Q_addAnnualSourceCheck_(add, annual, 'Dòng tiền huy động từ khách hàng', sum04('Dòng tiền huy động từ KH') / 1e9, 'Dòng tiền khách hàng', TOL_TY);
    FS99Q_addAnnualSourceCheck_(add, annual, 'Dòng tiền vốn CSH', sum04('Tổng dòng CSH vào dự án') / 1e9, 'Vốn CSH', TOL_TY);
    FS99Q_addAnnualSourceCheck_(add, annual, 'Dòng tiền vốn vay', sum04('Giải ngân vay') / 1e9, 'Giải ngân vay', TOL_TY);
    FS99Q_addAnnualSourceCheck_(add, annual, 'VAT phải nộp', sum04('VAT phải nộp') / 1e9, 'VAT phải nộp', TOL_TY);
    FS99Q_addAnnualSourceCheck_(add, annual, 'Thuế TNDN', sum04('Thuế TNDN') / 1e9, 'Thuế TNDN', TOL_TY);
    FS99Q_addAnnualSourceCheck_(add, annual, 'Lãi vay', sum04('Lãi vay vốn hóa') / 1e9, 'Lãi vay', TOL_TY);
    FS99Q_addAnnualSourceCheck_(add, annual, 'Trả gốc vay', sum04('Trả gốc') / 1e9, 'Trả gốc', TOL_TY);
    FS99Q_addAnnualSourceCheck_(add, annual, 'FCFF dự án', sum04('FCFF_TIPV') / 1e9, 'FCFF', TOL_TY);
    FS99Q_addAnnualSourceCheck_(add, annual, 'FCFE vốn CSH', sum04('FCFE khả dụng cho CSH') / 1e9, 'FCFE', TOL_TY);

    const totalIn = FS99Q_annualValue_(annual, 'TỔNG DÒNG TIỀN VÀO');
    const totalOut = FS99Q_annualValue_(annual, 'TỔNG DÒNG TIỀN RA');
    const net = FS99Q_annualValue_(annual, 'Dòng tiền thuần sau tài trợ');
    const equity = FS99Q_annualValue_(annual, '(-) Vốn CSH góp mới');
    const loan = FS99Q_annualValue_(annual, '(-) Vốn vay giải ngân');
    const principal = FS99Q_annualValue_(annual, '(+) Trả gốc vay');
    const interest = FS99Q_annualValue_(annual, '(+) Lãi vay');
    const fcffA = FS99Q_annualValue_(annual, 'FCFF dự án');
    const fcfeA = FS99Q_annualValue_(annual, 'FCFE vốn CSH');

    add('TẦNG 2 - 04A', 'Dòng tiền thuần sau tài trợ = Tổng vào - Tổng ra',
      net, totalIn - totalOut, 'Kiểm tra ngay trên Sheet 04A.',
      'Sửa hàm FS04A_buildAnnualFromSheet04_.', TOL_TY);
    add('TẦNG 2 - 04A', 'FCFF = Thuần sau tài trợ - CSH - Vay + Trả gốc + Lãi vay',
      fcffA, net - equity - loan + principal + interest,
      'Kiểm tra cầu nối đến FCFF.', 'Sửa công thức phần C Sheet 04A.', TOL_TY);
    add('TẦNG 2 - 04A', 'FCFE = Thuần sau tài trợ - CSH + Lãi vay',
      fcfeA, net - equity + interest,
      'Kiểm tra cầu nối đến FCFE.', 'Sửa công thức phần C Sheet 04A.', TOL_TY);
  } else {
    results.push(['TẦNG 2 - 04A', 'Tồn tại Sheet 04A. TH dòng tiền', 0, 1, -1, 'FAIL',
      'Không tìm thấy Sheet 04A.', 'Chạy 8.3 Pipeline - Tổng hợp theo năm.']);
  }

  // TẦNG 3 - SHEET 00
  const summary = FS99Q_readSummaryByLabel_(sh00);
  const srcRevenueTy = sum02('Dòng tiền huy động từ KH') / 1e9;
  const srcCostTy = (sum03('Tổng chi sau VAT') + sum04('Lãi vay vốn hóa')) / 1e9;
  const srcProfitTy = sum04('Lợi nhuận sau thuế') / 1e9;
  const srcCitTy = sum04('Thuế TNDN') / 1e9;
  const srcVatTy = sum04('VAT phải nộp') / 1e9;

  FS99Q_addSummaryCheck_(add, summary, 'Tổng doanh thu có VAT', srcRevenueTy,
    'Tổng doanh thu có VAT', 'Sửa module lập mục III Sheet 00.', TOL_TY);
  FS99Q_addSummaryCheck_(add, summary, 'Tổng chi phí có VAT', srcCostTy,
    'Tổng chi phí sau VAT + lãi vay', 'Xem Tầng 5 để xác định đúng cấu phần gây lệch.', TOL_TY);
  FS99Q_addSummaryCheck_(add, summary, 'Lợi nhuận sau thuế', srcProfitTy,
    'Lợi nhuận sau thuế', 'Sửa module lập mục III Sheet 00.', TOL_TY);
  FS99Q_addSummaryCheck_(add, summary, 'Tổng Thuế TNDN', srcCitTy,
    'Thuế TNDN', 'Sửa module lập mục III Sheet 00.', TOL_TY);
  FS99Q_addSummaryCheck_(add, summary, 'Tổng VAT phải nộp', srcVatTy,
    'VAT phải nộp', 'Sửa module lập mục III Sheet 00.', TOL_TY);

  // TẦNG 4 - MỤC 14
  const summaryRevenue = FS99Q_summaryValue_(summary, 'Tổng doanh thu có VAT');
  const summaryCost = FS99Q_summaryValue_(summary, 'Tổng chi phí có VAT');
  const summaryProfit = FS99Q_summaryValue_(summary, 'Lợi nhuận sau thuế');
  const summaryCit = FS99Q_summaryValue_(summary, 'Tổng Thuế TNDN');
  const summaryVat = FS99Q_summaryValue_(summary, 'Tổng VAT phải nộp');
  const scopeShown = FS99Q_summaryValue_(summary, 'Chênh lệch đối chiếu phạm vi');
  const scopeCalc = summaryRevenue - summaryCost - summaryProfit - summaryCit - summaryVat;

  add('TẦNG 4 - PHƯƠNG TRÌNH', 'Mục 14 = Doanh thu - Chi phí - LNST - TNDN - VAT',
    scopeShown, scopeCalc,
    'Chỉ kiểm tra công thức hiển thị mục 14; giá trị không bắt buộc bằng 0.',
    'Sửa hàm lập mục 14 Sheet 00.', TOL_TY);

  // TẦNG 5 - PHÂN RÃ CHÍNH XÁC TỔNG CHI PHÍ SHEET 00
  const sellingRate = FS99Q_getCommonCostVatRate_(tech, ['Chi phí bán hàng'], 0);
  const operatingRate = FS99Q_getCommonCostVatRate_(tech, ['Chi phí vận hành', 'Chi phí vận hành thuê'], sellingRate);
  const maintenanceRate = FS99Q_getCommonCostVatRate_(tech, ['Chi phí bảo trì', 'Chi phí bảo hành, bảo trì'], operatingRate);

  const srcSellingTy = sum03('Chi phí bán hàng trước VAT') * (1 + sellingRate) / 1e9;
  const srcOperatingTy = sum03('Chi phí vận hành thuê trước VAT') * (1 + operatingRate) / 1e9;
  const srcMaintenanceTy = sum03('Chi phí bảo trì trước VAT') * (1 + maintenanceRate) / 1e9;
  const srcInvestmentTy = srcCostTy - srcSellingTy - srcOperatingTy - srcMaintenanceTy;

  const shownInvestmentTy = FS99Q_summaryValue_(summary, 'Tổng vốn đầu tư dự án');
  const shownSellingTy = FS99Q_summaryValue_(summary, 'Chi phí bán hàng');
  const shownOperatingTy = FS99Q_summaryValue_(summary, 'Chi phí vận hành');
  const shownMaintenanceTy = FS99Q_summaryValue_(summary, 'Chi phí bảo trì');
  const shownDetailTotalTy = shownInvestmentTy + shownSellingTy + shownOperatingTy + shownMaintenanceTy;

  add('TẦNG 5 - PHÂN RÃ CHI PHÍ', '2.1 Tổng vốn đầu tư dự án = nguồn còn lại sau loại chi phí hoạt động',
    shownInvestmentTy, srcInvestmentTy,
    'Nguồn = Tổng chi sau VAT Sheet 03 + lãi vay - bán hàng - vận hành - bảo trì.',
    'Sửa công thức/nguồn của chỉ tiêu Tổng vốn đầu tư dự án trong hàm gốc FS_lapSheet00.', TOL_TY);
  add('TẦNG 5 - PHÂN RÃ CHI PHÍ', 'Chi phí bán hàng Sheet 00 = nguồn sau VAT',
    shownSellingTy, srcSellingTy,
    'VAT theo cấu hình Sheet 01. Kỹ thuật: ' + (sellingRate * 100) + '%.',
    'Sửa phần tính Chi phí bán hàng trong ZZZZZZZZ_SummaryRevenueCheckFix.js.', TOL_TY);
  add('TẦNG 5 - PHÂN RÃ CHI PHÍ', 'Chi phí vận hành Sheet 00 = nguồn sau VAT',
    shownOperatingTy, srcOperatingTy,
    'VAT theo cấu hình Sheet 01. Kỹ thuật: ' + (operatingRate * 100) + '%.',
    'Sửa phần tính Chi phí vận hành trong ZZZZZZZZ_SummaryRevenueCheckFix.js.', TOL_TY);
  add('TẦNG 5 - PHÂN RÃ CHI PHÍ', 'Chi phí bảo trì Sheet 00 = nguồn sau VAT',
    shownMaintenanceTy, srcMaintenanceTy,
    'VAT theo cấu hình Sheet 01. Kỹ thuật: ' + (maintenanceRate * 100) + '%.',
    'Sửa phần tính Chi phí bảo trì trong ZZZZZZZZ_SummaryRevenueCheckFix.js.', TOL_TY);
  add('TẦNG 5 - PHÂN RÃ CHI PHÍ', 'Tổng các dòng chi tiết Sheet 00 = Tổng chi phí có VAT Sheet 00',
    shownDetailTotalTy, summaryCost,
    'Kiểm tra phép cộng các dòng 2.1 và các khoản chi phí có phát sinh.',
    'Sửa công thức Tổng chi phí có VAT tại Sheet 00.', TOL_TY);
  add('TẦNG 5 - PHÂN RÃ CHI PHÍ', 'Tổng nguồn phân rã = nguồn Tổng chi phí có VAT',
    srcInvestmentTy + srcSellingTy + srcOperatingTy + srcMaintenanceTy, srcCostTy,
    'Phép kiểm soát nội bộ của bảng phân rã.',
    'Kiểm tra logic phân rã nguồn trong module audit.', TOL_TY);

  const failRows = results.filter(r => r[5] === 'FAIL');
  let rootCause;
  if (failRows.length) {
    const first = failRows[0];
    const costFails = failRows.filter(r => r[0] === 'TẦNG 5 - PHÂN RÃ CHI PHÍ');
    rootCause = costFails.length
      ? 'Sai lệch Tổng chi phí được phân rã tại Tầng 5. Dòng FAIL đầu tiên: ' + costFails[0][1] + '. ' + costFails[0][7]
      : 'Lỗi đầu tiên nằm tại ' + first[0] + ': ' + first[1] + '. ' + first[7];
    results.push(['KẾT LUẬN', 'Vị trí sửa đầu tiên', first[2], first[3], first[4], 'FAIL',
      'Ưu tiên kết luận phân rã Tầng 5 nếu sai lệch liên quan Tổng chi phí.', rootCause]);
  } else {
    rootCause = 'Các nguồn chi tiết, Sheet 04A, Sheet 00 và bảng phân rã chi phí đều khớp.';
    results.push(['KẾT LUẬN', 'Kết quả audit phân tầng', 0, 0, 0, 'PASS', '', rootCause]);
  }

  FS99Q_writeResults_(shC, results);

  const msg = failRows.length
    ? 'AUDIT FAILED: ' + failRows.length + ' sai lệch.\n\n' + rootCause + '\n\nXem Sheet 99. Checks từ dòng 50.'
    : 'AUDIT PASS: các nguồn chi tiết, Sheet 04A và Sheet 00 đều khớp.';
  SpreadsheetApp.getUi().alert(msg);
  return { pass: failRows.length === 0, failures: failRows, rootCause };
}

function FS99Q_addAnnualSourceCheck_(add, annual, label, sourceValue, shortName, tol) {
  const actual = FS99Q_annualValue_(annual, label);
  add('TẦNG 2 - 04A', label + ' = nguồn Sheet 04', actual, sourceValue,
    'Đơn vị tỷ đồng.', 'Sửa hàm tổng hợp theo năm tại dòng ' + shortName + ' của Sheet 04A.', tol);
}

function FS99Q_addSummaryCheck_(add, summary, label, sourceValue, shortName, fix, tol) {
  const actual = FS99Q_summaryValue_(summary, label);
  add('TẦNG 3 - SHEET 00', label + ' = nguồn chi tiết', actual, sourceValue,
    shortName + ' / 1 tỷ.', fix, tol);
}

function FS99Q_readAnnualTotals_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return {};
  const values = sheet.getRange(1, 1, lastRow, Math.min(3, sheet.getLastColumn())).getValues();
  const out = {};
  for (let i = 0; i < values.length; i++) {
    const label = FS99Q_key_(values[i][1]);
    if (label) out[label] = FS99Q_num_(values[i][2]);
  }
  return out;
}

function FS99Q_annualValue_(annual, label) {
  const key = FS99Q_key_(label);
  if (key in annual) return FS99Q_num_(annual[key]);
  const matches = Object.keys(annual).filter(k => k.indexOf(key) === 0);
  if (matches.length === 1) return FS99Q_num_(annual[matches[0]]);
  if (matches.length > 1) throw new Error('Có nhiều dòng Sheet 04A cùng bắt đầu bằng: ' + label);
  throw new Error('Không tìm thấy dòng trên Sheet 04A: ' + label);
}

function FS99Q_readSummaryByLabel_(sheet) {
  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(1, 2, lastRow, 3).getValues(); // B:D
  const out = {};
  for (let i = 0; i < values.length; i++) {
    const label = FS99Q_key_(values[i][0]);
    if (label) out[label] = { row: i + 1, value: FS99Q_num_(values[i][2]) };
  }
  return out;
}

function FS99Q_summaryValue_(summary, label) {
  const item = summary[FS99Q_key_(label)];
  return item ? FS99Q_num_(item.value) : 0;
}

function FS99Q_detectHeader_(sheet, requiredNames) {
  const maxRows = Math.min(6, sheet.getLastRow());
  let best = null;
  for (let row = 1; row <= maxRows; row++) {
    const raw = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    const headers = raw.map(FS99Q_key_);
    const score = requiredNames.reduce((s, name) => s + (headers.indexOf(FS99Q_key_(name)) >= 0 ? 1 : 0), 0);
    if (!best || score > best.score) best = { headerRow: row, headers, score };
  }
  if (!best || best.score === 0) throw new Error('Không nhận diện được dòng tiêu đề tại sheet ' + sheet.getName() + '.');
  return best;
}

function FS99Q_sumNamed_(sheet, meta, name) {
  const col = meta.headers.indexOf(FS99Q_key_(name)) + 1;
  if (col < 1) return 0;
  return FS99Q_sumCol_(sheet, meta.headerRow + 1, col);
}

function FS99Q_sumCol_(sheet, startRow, col) {
  if (!col || col < 1 || sheet.getLastRow() < startRow) return 0;
  return sheet.getRange(startRow, col, sheet.getLastRow() - startRow + 1, 1)
    .getValues().reduce((s, r) => s + FS99Q_num_(r[0]), 0);
}

function FS99Q_getCommonCostVatRate_(tech, names, fallback) {
  const values = tech.getRange(1, 1, tech.getLastRow(), Math.max(3, Math.min(tech.getLastColumn(), 8))).getValues();
  const targets = names.map(FS99Q_key_);
  let inBlock = false;
  for (let r = 0; r < values.length; r++) {
    const firstRaw = String(values[r][0] || '').trim();
    const first = FS99Q_key_(firstRaw);
    if (first === 'chi phi chung' || first === 'chi_phi_chung') {
      inBlock = true;
      continue;
    }
    if (!inBlock) continue;
    if (r > 0 && /^[A-Z0-9_]{4,}$/.test(firstRaw) && first !== 'chi phi chung') break;
    if (targets.indexOf(first) >= 0) return FS99Q_rate_(values[r][2]);
  }
  return FS99Q_rate_(fallback);
}

function FS99Q_rate_(value) {
  let n = Number(value);
  if (!isFinite(n)) n = 0;
  if (Math.abs(n) > 1) n = n / 100;
  return Math.max(0, n);
}

function FS99Q_writeResults_(sheet, results) {
  const startRow = 50;
  const cols = 8;
  const clearRows = Math.max(1, sheet.getMaxRows() - startRow + 1);
  sheet.getRange(startRow, 1, clearRows, cols).clearContent().clearFormat();
  sheet.getRange(startRow, 1)
    .setValue('AUDIT RECONCILIATION PHÂN TẦNG - XÁC ĐỊNH ĐÚNG MODULE CẦN SỬA')
    .setFontWeight('bold').setFontSize(14);
  sheet.getRange(startRow + 1, 1, 1, cols)
    .setValues([['Tầng / Nhóm','Phép kiểm tra','Nguồn A','Nguồn B','Chênh lệch','Trạng thái','Ghi chú','Kết luận / module cần sửa']])
    .setFontWeight('bold').setBackground('#d9ead3');
  sheet.getRange(startRow + 2, 1, results.length, cols).setValues(results);
  sheet.getRange(startRow + 2, 3, results.length, 3).setNumberFormat('#,##0.000');
  sheet.getRange(startRow + 2, 6, results.length, 1).setFontWeight('bold');
  sheet.getRange(startRow + 2, 7, results.length, 2).setWrap(true);
  sheet.autoResizeColumns(1, 7);
  sheet.setColumnWidth(8, 520);

  const statuses = sheet.getRange(startRow + 2, 6, results.length, 1).getValues();
  for (let i = 0; i < statuses.length; i++) {
    const row = startRow + 2 + i;
    const status = statuses[i][0];
    if (status === 'FAIL') sheet.getRange(row, 1, 1, cols).setBackground('#f4cccc');
    if (status === 'INFO') sheet.getRange(row, 1, 1, cols).setBackground('#fff2cc');
  }
}

function FS99Q_num_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  const s = String(value == null ? '' : value).trim().replace(/\s/g, '').replace(/\./g, '').replace(',', '.');
  const n = Number(s);
  return isFinite(n) ? n : 0;
}

function FS99Q_key_(value) {
  return String(value || '').toLowerCase().normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/[^a-z0-9_]+/g, ' ')
    .trim();
}