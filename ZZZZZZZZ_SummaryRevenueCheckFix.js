/*************************************************
 * ZZZZZZZZ_SummaryRevenueCheckFix.js
 *
 * Nguyên tắc:
 * - Chỉ tiêu 1 trên Sheet 00 luôn lấy trực tiếp từ Sheet 02.
 * - Chi phí bán hàng, vận hành và bảo trì tại mục III đều trình bày sau VAT.
 * - VAT từng khoản chi lấy theo cấu hình tại Sheet 01. Kỹ thuật.
 * - Các khoản 2.3 và 2.4 nằm ngay sau khoản 2.2.
 * - Hàm chạy lặp lại không được nhân đôi dòng hoặc làm mất Thuế TNDN/VAT.
 *************************************************/

const FSZZZZZZZZ_BASE_LAP_SHEET00_ = FS_lapSheet00;

FS_lapSheet00 = function() {
  FSZZZZZZZZ_BASE_LAP_SHEET00_();
  FSZZZZZZZZ_restoreSummaryRevenueAndCheck_();
};

function FSZZZZZZZZ_suaTongHopDoanhThuVaKiemTra() {
  FSZZZZZZZZ_restoreSummaryRevenueAndCheck_();
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Đã sắp xếp lại mục chi phí và cập nhật công thức Sheet 00.');
}

function FSZZZZZZZZ_restoreSummaryRevenueAndCheck_() {
  const ss = SpreadsheetApp.getActive();
  const sh00 = ss.getSheetByName('00. Tổng hợp');
  const tech = ss.getSheetByName('01. Kỹ thuật');
  const sh02 = ss.getSheetByName('02. Doanh thu');
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  if (!sh00 || !tech || !sh02 || !sh03) {
    throw new Error('Thiếu Sheet 00. Tổng hợp, 01. Kỹ thuật, 02. Doanh thu hoặc 03. Chi phí & Vốn.');
  }

  const headers02 = sh02.getRange(1, 1, 1, sh02.getLastColumn()).getDisplayValues()[0];
  const headers03 = sh03.getRange(1, 1, 1, sh03.getLastColumn()).getDisplayValues()[0];

  const cashCol = FSZZZZZZZZ_findHeader_(headers02, ['Dòng tiền huy động từ KH','Dòng tiền huy động từ khách hàng']);
  const sellingCol = FSZZZZZZZZ_findHeader_(headers03, ['Chi phí bán hàng trước VAT']);
  const opCol = FSZZZZZZZZ_findHeader_(headers03, ['Chi phí vận hành thuê trước VAT','Chi phí vận hành trước VAT']);
  const maintCol = FSZZZZZZZZ_findHeader_(headers03, ['Chi phí bảo trì trước VAT']);
  if (cashCol < 1) throw new Error('Không tìm thấy cột Dòng tiền huy động từ KH trên Sheet 02.');

  const sellingVatRate = FSZZZZZZZZ_getCommonCostVatRate_(tech, ['Chi phí bán hàng'], 0);
  const opVatRate = FSZZZZZZZZ_getCommonCostVatRate_(tech, ['Chi phí vận hành','Chi phí vận hành thuê'], sellingVatRate);
  const maintVatRate = FSZZZZZZZZ_getCommonCostVatRate_(tech, ['Chi phí bảo trì','Chi phí bảo hành, bảo trì'], opVatRate);

  const cashLetter = FSZZZZZZZZ_colLetter_(cashCol);
  sh00.getRange('D25').setFormula(`=SUM('02. Doanh thu'!${cashLetter}2:${cashLetter})/1000000000`);

  // Chỉ dịch khối chỉ tiêu gốc một lần. Nếu dòng 33 đã là Chi phí vận hành,
  // hàm đang được gọi lặp lại trong cùng pipeline và không được copy lần hai.
  const alreadyExpanded = FSZZZZZZZZ_norm_(sh00.getRange('B33').getDisplayValue()).indexOf('chi phi van hanh') >= 0;
  if (!alreadyExpanded) {
    sh00.getRange('A33:E43').copyTo(
      sh00.getRange('A35:E45'),
      SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
      false
    );
  }

  const sellingAfterVat = FSZZZZZZZZ_sumColumn_(sh03, sellingCol, 2) * (1 + sellingVatRate) / 1e9;
  const opAfterVat = FSZZZZZZZZ_sumColumn_(sh03, opCol, 2) * (1 + opVatRate) / 1e9;
  const maintAfterVat = FSZZZZZZZZ_sumColumn_(sh03, maintCol, 2) * (1 + maintVatRate) / 1e9;

  // Dùng định dạng của dòng 2.2 làm chuẩn cho 2.3 và 2.4.
  sh00.getRange('A32:E32').copyTo(
    sh00.getRange('A33:E34'),
    SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
    false
  );

  sh00.getRange('A32:E34').setValues([
    ['2.2', 'Chi phí bán hàng (sau VAT)', 'tỷ đồng', sellingAfterVat, ''],
    ['2.3', 'Chi phí vận hành (sau VAT)', 'tỷ đồng', opAfterVat, ''],
    ['2.4', 'Chi phí bảo trì (sau VAT)', 'tỷ đồng', maintAfterVat, '']
  ]);
  sh00.getRange('E32:E34').clearContent();
  sh00.getRange('A32:E34')
    .setFontColor('#000000')
    .setBackground('#ffffff')
    .setFontWeight('normal')
    .setVerticalAlignment('middle');
  sh00.getRange('A32:A34').setHorizontalAlignment('center');
  sh00.getRange('C32:C34').setHorizontalAlignment('center');
  sh00.getRange('D32:D34')
    .setHorizontalAlignment('right')
    .setNumberFormat('#,##0.0;[Red]-#,##0.0');

  // Tổng chi phí có VAT = vốn đầu tư + ba khoản chi phí hoạt động sau VAT.
  sh00.getRange('D30').setFormula('=SUM(D31:D34)');

  // Khôi phục chắc chắn hai chỉ tiêu thuế sau khi dịch khối.
  sh00.getRange('A44:E45').setValues([
    ['12', 'Tổng Thuế TNDN', 'tỷ đồng', '', ''],
    ['13', 'Tổng VAT phải nộp', 'tỷ đồng', '', '']
  ]);
  sh00.getRange('D44').setFormula(`=SUM('04. Dòng tiền & Lợi nhuận'!L3:L)/1000000000`);
  sh00.getRange('D45').setFormula(`=SUM('04. Dòng tiền & Lợi nhuận'!K3:K)/1000000000`);

  const checkRow = 46;
  sh00.getRange(checkRow, 1, 1, 5).clearContent();
  sh00.getRange(checkRow, 1).setValue('14');
  sh00.getRange(checkRow, 2).setValue('Chênh lệch đối chiếu phạm vi');
  sh00.getRange(checkRow, 3).setValue('tỷ đồng');
  sh00.getRange(checkRow, 4).setFormula('=D25-(D30+D35+D44+D45)');
  sh00.getRange(checkRow, 5).clearContent();

  sh00.getRange('A35:E46')
    .setFontFamily('Times New Roman')
    .setFontSize(11)
    .setBorder(true, true, true, true, true, true)
    .setVerticalAlignment('middle');
  sh00.getRange('D30:D46').setNumberFormat('#,##0.0;[Red]-#,##0.0');
  sh00.getRange('A44:C46').setFontWeight('bold');
  sh00.getRange('D44:D46').setFontWeight('bold');
  sh00.setRowHeight(33, 24);
  sh00.setRowHeight(34, 24);
  sh00.setRowHeight(46, 24);

  SpreadsheetApp.flush();
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

function FSZZZZZZZZ_colLetter_(col) {
  let out = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    col = Math.floor((col - 1) / 26);
  }
  return out;
}
