/*************************************************
 * ZZZZZZZZ_SummaryRevenueCheckFix.js
 * Ổn định cố định mục III Sheet 00 sau khi hàm gốc lập báo cáo.
 *
 * Nguyên tắc:
 * - Mục 2 đã ghi "Tổng chi phí có VAT", nên các mục 2.1-2.4 không lặp lại "có VAT/sau VAT".
 * - Không chèn/copy lặp làm nhân đôi dòng.
 * - Dựng lại vùng A25:E46 theo bố cục cố định sau mỗi lần lập Sheet 00.
 * - Chi phí bán hàng, vận hành, bảo trì lấy theo header thực tế Sheet 03.
 * - VAT từng khoản lấy từ cấu hình Sheet 01. Kỹ thuật.
 *************************************************/

const FSZZZZZZZZ_BASE_LAP_SHEET00_ = FS_lapSheet00;

FS_lapSheet00 = function() {
  FSZZZZZZZZ_BASE_LAP_SHEET00_();
  FSZZZZZZZZ_rebuildStableSummary_();
};

function FSZZZZZZZZ_suaTongHopDoanhThuVaKiemTra() {
  FSZZZZZZZZ_BASE_LAP_SHEET00_();
  FSZZZZZZZZ_rebuildStableSummary_();
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Đã dựng lại ổn định mục III Sheet 00.');
}

function FSZZZZZZZZ_rebuildStableSummary_() {
  const ss = SpreadsheetApp.getActive();
  const sh00 = ss.getSheetByName('00. Tổng hợp');
  const tech = ss.getSheetByName('01. Kỹ thuật');
  const sh02 = ss.getSheetByName('02. Doanh thu');
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');
  if (!sh00 || !tech || !sh02 || !sh03 || !sh04) {
    throw new Error('Thiếu Sheet 00, 01. Kỹ thuật, 02, 03 hoặc 04.');
  }

  const h02 = FSZZZZZZZZ_detectHeader_(sh02, ['Dòng tiền huy động từ KH', 'Tổng doanh thu trước VAT']);
  const h03 = FSZZZZZZZZ_detectHeader_(sh03, ['Chi phí bán hàng trước VAT', 'Tổng chi sau VAT']);
  const h04 = FSZZZZZZZZ_detectHeader_(sh04, ['Lợi nhuận sau thuế', 'Thuế TNDN', 'VAT phải nộp']);

  const cashCol = FSZZZZZZZZ_findHeader_(h02.headers, ['Dòng tiền huy động từ KH', 'Dòng tiền huy động từ khách hàng']);
  const sellingCol = FSZZZZZZZZ_findHeader_(h03.headers, ['Chi phí bán hàng trước VAT']);
  const opCol = FSZZZZZZZZ_findHeader_(h03.headers, ['Chi phí vận hành thuê trước VAT', 'Chi phí vận hành trước VAT']);
  const maintCol = FSZZZZZZZZ_findHeader_(h03.headers, ['Chi phí bảo trì trước VAT']);
  const totalAfterVatCol = FSZZZZZZZZ_findHeader_(h03.headers, ['Tổng chi sau VAT']);
  const profitAfterTaxCol = FSZZZZZZZZ_findHeader_(h04.headers, ['Lợi nhuận sau thuế']);
  const citCol = FSZZZZZZZZ_findHeader_(h04.headers, ['Thuế TNDN']);
  const vatPayCol = FSZZZZZZZZ_findHeader_(h04.headers, ['VAT phải nộp']);
  const interestCol = FSZZZZZZZZ_findHeader_(h04.headers, ['Lãi vay vốn hóa']);

  const required = [cashCol, sellingCol, opCol, maintCol, totalAfterVatCol, profitAfterTaxCol, citCol, vatPayCol, interestCol];
  if (required.some(c => c < 1)) {
    throw new Error('Không tìm đủ cột nguồn để lập mục III Sheet 00. Kiểm tra header Sheet 02, 03 và 04.');
  }

  const sellingVatRate = FSZZZZZZZZ_getCommonCostVatRate_(tech, ['Chi phí bán hàng'], 0);
  const opVatRate = FSZZZZZZZZ_getCommonCostVatRate_(tech, ['Chi phí vận hành', 'Chi phí vận hành thuê'], sellingVatRate);
  const maintVatRate = FSZZZZZZZZ_getCommonCostVatRate_(tech, ['Chi phí bảo trì', 'Chi phí bảo hành, bảo trì'], opVatRate);

  const sellingAfterVat = FSZZZZZZZZ_sumColumn_(sh03, sellingCol, h03.dataStartRow) * (1 + sellingVatRate) / 1e9;
  const opAfterVat = FSZZZZZZZZ_sumColumn_(sh03, opCol, h03.dataStartRow) * (1 + opVatRate) / 1e9;
  const maintAfterVat = FSZZZZZZZZ_sumColumn_(sh03, maintCol, h03.dataStartRow) * (1 + maintVatRate) / 1e9;

  // Lưu công thức/giá trị D31 do hàm gốc vừa tạo trước khi ghi đè vùng A30:E34.
  const originalD31Formula = sh00.getRange('D31').getFormula();
  const originalD31Value = sh00.getRange('D31').getValue();

  // Chuyển nguyên khối từ LNST đến VAT xuống 2 dòng để dành chỗ cho 2.3 và 2.4.
  sh00.getRange('A33:E43').copyTo(
    sh00.getRange('A35:E45'),
    SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
    false
  );

  // Dựng cố định nhóm chi phí.
  sh00.getRange('A30:E34').setValues([
    ['2', 'Tổng chi phí có VAT', 'tỷ đồng', '', ''],
    ['2.1', 'Tổng vốn đầu tư dự án', 'tỷ đồng', '', ''],
    ['2.2', 'Chi phí bán hàng', 'tỷ đồng', sellingAfterVat, ''],
    ['2.3', 'Chi phí vận hành', 'tỷ đồng', opAfterVat, ''],
    ['2.4', 'Chi phí bảo trì', 'tỷ đồng', maintAfterVat, '']
  ]);

  // Khôi phục D31 sau khi setValues đã xóa công thức/giá trị cũ.
  if (originalD31Formula) {
    sh00.getRange('D31').setFormula(originalD31Formula);
  } else {
    sh00.getRange('D31').setValue(originalD31Value);
  }
  sh00.getRange('D30').setFormula('=SUM(D31:D34)');

  const citTy = FSZZZZZZZZ_sumColumn_(sh04, citCol, h04.dataStartRow) / 1e9;
  const vatPayTy = FSZZZZZZZZ_sumColumn_(sh04, vatPayCol, h04.dataStartRow) / 1e9;
  sh00.getRange('A44:E45').setValues([
    ['12', 'Tổng Thuế TNDN', 'tỷ đồng', citTy, ''],
    ['13', 'Tổng VAT phải nộp', 'tỷ đồng', vatPayTy, '']
  ]);

  const revenueTy = FSZZZZZZZZ_sumColumn_(sh02, cashCol, h02.dataStartRow) / 1e9;
  const totalCostTy = (
    FSZZZZZZZZ_sumColumn_(sh03, totalAfterVatCol, h03.dataStartRow) +
    FSZZZZZZZZ_sumColumn_(sh04, interestCol, h04.dataStartRow)
  ) / 1e9;
  const profitAfterTaxTy = FSZZZZZZZZ_sumColumn_(sh04, profitAfterTaxCol, h04.dataStartRow) / 1e9;
  const scopeDifferenceTy = revenueTy - totalCostTy - profitAfterTaxTy - citTy - vatPayTy;

  sh00.getRange('A46:E46').setValues([
    ['14', 'Chênh lệch đối chiếu phạm vi', 'tỷ đồng', scopeDifferenceTy, '']
  ]);

  sh00.getRange('A25:E46')
    .setFontFamily('Times New Roman')
    .setFontSize(11)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true);

  sh00.getRange('A25:A46').setHorizontalAlignment('center');
  sh00.getRange('B25:B46').setHorizontalAlignment('left');
  sh00.getRange('C25:C46').setHorizontalAlignment('center');
  sh00.getRange('D25:D46').setHorizontalAlignment('right');
  sh00.getRange('E25:E46').setHorizontalAlignment('center');

  const moneyRows = [25,26,27,28,29,30,31,32,33,34,35,36,39,42,43,44,45,46];
  moneyRows.forEach(r => sh00.getRange(r, 4).setNumberFormat('#,##0.0;[Red]-#,##0.0'));
  [38,41].forEach(r => sh00.getRange(r, 4).setNumberFormat('0.0'));
  [37,40].forEach(r => sh00.getRange(r, 4).setNumberFormat('0.0%'));

  sh00.getRange('A31:E34')
    .setBackground('#ffffff')
    .setFontColor('#000000')
    .setFontWeight('normal');

  [25,30,35,42,43,44,45,46].forEach(r => sh00.getRange(r, 1, 1, 4).setFontWeight('bold'));
  sh00.getRange('E32:E34').clearContent();
  sh00.getRange('E44:E46').clearContent();
  for (let r = 25; r <= 46; r++) sh00.setRowHeight(r, 24);

  SpreadsheetApp.flush();
}

function FSZZZZZZZZ_detectHeader_(sheet, requiredNames) {
  const maxRows = Math.min(5, sheet.getLastRow());
  const lastCol = sheet.getLastColumn();
  for (let r = 1; r <= maxRows; r++) {
    const headers = sheet.getRange(r, 1, 1, lastCol).getDisplayValues()[0];
    const hits = requiredNames.filter(name => FSZZZZZZZZ_findHeader_(headers, [name]) > 0).length;
    if (hits > 0) return { headerRow: r, dataStartRow: r + 1, headers };
  }
  throw new Error('Không xác định được dòng tiêu đề tại sheet ' + sheet.getName());
}

function FSZZZZZZZZ_sumColumn_(sheet, col, startRow) {
  if (!col || col < 1 || sheet.getLastRow() < startRow) return 0;
  return sheet.getRange(startRow, col, sheet.getLastRow() - startRow + 1, 1)
    .getValues()
    .reduce((sum, row) => sum + (Number(row[0]) || 0), 0);
}

function FSZZZZZZZZ_getCommonCostVatRate_(tech, names, fallback) {
  const lastRow = tech.getLastRow();
  const lastCol = Math.max(3, Math.min(tech.getLastColumn(), 8));
  if (lastRow < 1) return FSZZZZZZZZ_rate_(fallback);

  const values = tech.getRange(1, 1, lastRow, lastCol).getValues();
  const targets = names.map(FSZZZZZZZZ_norm_);
  let inBlock = false;
  for (let r = 0; r < values.length; r++) {
    const firstRaw = String(values[r][0] || '').trim();
    const first = FSZZZZZZZZ_norm_(firstRaw);
    if (first === 'chi_phi_chung' || first === 'chi phi chung') {
      inBlock = true;
      continue;
    }
    if (!inBlock) continue;
    if (r > 0 && /^[A-Z0-9_]{4,}$/.test(firstRaw) && first !== 'chi_phi_chung') break;
    if (targets.indexOf(first) >= 0) return FSZZZZZZZZ_rate_(values[r][2]);
  }
  return FSZZZZZZZZ_rate_(fallback);
}

function FSZZZZZZZZ_rate_(value) {
  let n = Number(value);
  if (!isFinite(n)) n = 0;
  if (Math.abs(n) > 1) n = n / 100;
  return Math.max(0, n);
}

function FSZZZZZZZZ_findHeader_(headers, names) {
  const normalized = headers.map(FSZZZZZZZZ_norm_);
  for (const name of names) {
    const i = normalized.indexOf(FSZZZZZZZZ_norm_(name));
    if (i >= 0) return i + 1;
  }
  return -1;
}

function FSZZZZZZZZ_norm_(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}
