/*************************************************
 * 99Q_AuditReconciliation.js
 * Audit phân rã, tự nhận diện dòng tiêu đề thực tế của từng sheet.
 *************************************************/

function FS99P_buoc5_KiemTraNhanh() {
  return FS99Q_chayAuditReconciliation();
}

function FS99Q_chayAuditReconciliation() {
  const ss = SpreadsheetApp.getActive();
  const sh02 = ss.getSheetByName('02. Doanh thu');
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');
  const sh00 = ss.getSheetByName('00. Tổng hợp');
  let shC = ss.getSheetByName('99. Checks');
  if (!sh02 || !sh03 || !sh04 || !sh00) throw new Error('Thiếu Sheet 00, 02, 03 hoặc 04 để audit.');
  if (!shC) shC = ss.insertSheet('99. Checks');

  SpreadsheetApp.flush();

  const tolDong = 1;
  const tolTy = 0.001;
  const results = [];
  const add = (group, test, a, b, note, fix, tol) => {
    const d = FS99Q_num_(a) - FS99Q_num_(b);
    const limit = tol == null ? tolDong : tol;
    results.push([
      group, test, FS99Q_num_(a), FS99Q_num_(b), d,
      Math.abs(d) <= limit ? 'PASS' : 'FAIL',
      note || '',
      Math.abs(d) <= limit ? '' : (fix || 'Kiểm tra nguồn/công thức của chỉ tiêu này.')
    ]);
  };

  // Sheet 02 dùng tiêu đề dòng 1; Sheet 03 thường dùng dòng 2; Sheet 04 dùng dòng 2.
  // Hàm tự dò để không trả về 0 khi cấu trúc tiêu đề khác nhau.
  const meta02 = FS99Q_detectHeader_(sh02, ['Tháng số', 'Dòng tiền huy động từ KH', 'Thuế TNDN tạm tính']);
  const meta03 = FS99Q_detectHeader_(sh03, ['Tháng số', 'Tổng chi sau VAT', 'Thuế TNDN']);
  const meta04 = FS99Q_detectHeader_(sh04, ['Tháng số', 'FCFF_TIPV', 'Thuế TNDN']);

  const sum02 = n => FS99Q_sumCol_(sh02, meta02.headerRow + 1, FS99Q_col_(meta02.headers, n));
  const sum03 = n => FS99Q_sumCol_(sh03, meta03.headerRow + 1, FS99Q_col_(meta03.headers, n));
  const sum04 = n => FS99Q_sumCol_(sh04, meta04.headerRow + 1, FS99Q_col_(meta04.headers, n));

  add('CẤU TRÚC', 'Nhận diện dòng tiêu đề Sheet 02', meta02.headerRow, 1,
    'Dòng tiêu đề được phát hiện tự động.', 'Kiểm tra lại bố cục Sheet 02.', 0);
  add('CẤU TRÚC', 'Nhận diện dòng tiêu đề Sheet 03', meta03.headerRow, 2,
    'Dòng tiêu đề được phát hiện tự động.', 'Kiểm tra lại bố cục Sheet 03.', 0);
  add('CẤU TRÚC', 'Nhận diện dòng tiêu đề Sheet 04', meta04.headerRow, 2,
    'Dòng tiêu đề được phát hiện tự động.', 'Kiểm tra lại bố cục Sheet 04.', 0);

  add('DOANH THU', 'Sheet 02: Tổng doanh thu trước VAT = bán + thuê',
    sum02('Tổng doanh thu trước VAT'),
    sum02('Doanh thu bán trước VAT') + sum02('Doanh thu thuê trước VAT'));
  add('DOANH THU', 'Dòng tiền huy động KH Sheet 02 = Sheet 03',
    sum02('Dòng tiền huy động từ KH'), sum03('Dòng tiền huy động từ KH'));

  const costParts = [
    'Chi XD/TB/khác trước VAT', 'Chi GPMB trước VAT', 'Tiền SDĐ/thuê đất trước VAT',
    'Chi HTKT trước VAT', 'Chi phí bán hàng trước VAT', 'Chi phí dự phòng trước VAT',
    'Chi phí vận hành thuê trước VAT', 'Chi phí bảo trì trước VAT'
  ];
  const parts03 = costParts.reduce((s, n) => s + sum03(n), 0);
  add('CHI PHÍ', 'Sheet 03: Tổng chi trước VAT = tổng 8 cấu phần',
    sum03('Tổng chi trước VAT'), parts03,
    'XD + GPMB + đất + HTKT + bán hàng + dự phòng + vận hành + bảo trì');
  add('CHI PHÍ', 'Sheet 03: Tổng chi sau VAT = Tổng chi trước VAT + VAT đầu vào',
    sum03('Tổng chi sau VAT'), sum03('Tổng chi trước VAT') + sum03('VAT đầu vào'));

  add('THUẾ', 'Thuế TNDN Sheet 02 = Sheet 03', sum02('Thuế TNDN tạm tính'), sum03('Thuế TNDN'));
  add('THUẾ', 'Thuế TNDN Sheet 03 = Sheet 04', sum03('Thuế TNDN'), sum04('Thuế TNDN'));

  const inflow = sum04('Dòng tiền huy động từ KH') + sum04('Tổng dòng CSH vào dự án') + sum04('Giải ngân vay');
  const outflow = parts03 + sum03('VAT đầu vào') + sum04('VAT phải nộp') + sum04('Thuế TNDN') +
    sum04('Lãi vay vốn hóa') + sum04('Trả gốc');
  const netAfterFinancing = inflow - outflow;

  add('DÒNG TIỀN', 'FCFF Sheet 04 = Dòng tiền trước tài trợ',
    sum04('FCFF_TIPV'), sum04('Dòng tiền trước tài trợ'));
  add('DÒNG TIỀN', 'FCFE = FCFF + Giải ngân vay - Trả gốc',
    sum04('FCFE khả dụng cho CSH'),
    sum04('FCFF_TIPV') + sum04('Giải ngân vay') - sum04('Trả gốc'));
  add('DÒNG TIỀN', 'Dòng tiền thuần sau tài trợ = Tổng vào - Tổng ra',
    sum04('Dòng tiền thuần sau tài trợ'), netAfterFinancing);
  add('DÒNG TIỀN', 'FCFE = Dòng tiền thuần - CSH góp + Lãi vay',
    sum04('FCFE khả dụng cho CSH'),
    netAfterFinancing - sum04('Tổng dòng CSH vào dự án') + sum04('Lãi vay vốn hóa'));
  add('DÒNG TIỀN', 'FCFF = Dòng tiền thuần - CSH - Vay + Trả gốc + Lãi vay',
    sum04('FCFF_TIPV'),
    netAfterFinancing - sum04('Tổng dòng CSH vào dự án') - sum04('Giải ngân vay') +
      sum04('Trả gốc') + sum04('Lãi vay vốn hóa'));

  const srcRevenueTy = sum02('Dòng tiền huy động từ KH') / 1e9;
  const srcCostTy = (sum03('Tổng chi sau VAT') + sum04('Lãi vay vốn hóa')) / 1e9;
  const srcProfitAfterTaxTy = sum04('Lợi nhuận sau thuế') / 1e9;
  const srcCitTy = sum03('Thuế TNDN') / 1e9;
  const srcVatPayTy = sum04('VAT phải nộp') / 1e9;

  const d25 = FS99Q_num_(sh00.getRange('D25').getValue());
  const d30 = FS99Q_num_(sh00.getRange('D30').getValue());
  const d35 = FS99Q_num_(sh00.getRange('D35').getValue());
  const d44 = FS99Q_num_(sh00.getRange('D44').getValue());
  const d45 = FS99Q_num_(sh00.getRange('D45').getValue());

  add('TỔNG HỢP - NGUỒN', 'D25 Tổng doanh thu có VAT = Sheet 02', d25, srcRevenueTy,
    'Tổng Dòng tiền huy động từ KH / 1 tỷ.', 'Sửa D25 hoặc hàm lập Sheet 00.', tolTy);
  add('TỔNG HỢP - NGUỒN', 'D30 Tổng chi phí có VAT = chi phí sau VAT + lãi vay', d30, srcCostTy,
    'Tổng chi sau VAT Sheet 03 + lãi vay vốn hóa Sheet 04.', 'Sửa D30 hoặc D31:D34.', tolTy);
  add('TỔNG HỢP - NGUỒN', 'D35 Lợi nhuận sau thuế = Sheet 04', d35, srcProfitAfterTaxTy,
    'Tổng Lợi nhuận sau thuế / 1 tỷ.', 'Sửa D35.', tolTy);
  add('TỔNG HỢP - NGUỒN', 'D44 Tổng Thuế TNDN = Sheet 03', d44, srcCitTy,
    'Tổng Thuế TNDN / 1 tỷ.', 'Sửa D44.', tolTy);
  add('TỔNG HỢP - NGUỒN', 'D45 Tổng VAT phải nộp = Sheet 04', d45, srcVatPayTy,
    'Tổng VAT phải nộp / 1 tỷ.', 'Sửa D45.', tolTy);

  const summaryRight = d30 + d35 + d44 + d45;
  const summaryDiff = d25 - summaryRight;
  const componentDiffs = [
    ['D25', d25 - srcRevenueTy], ['D30', d30 - srcCostTy], ['D35', d35 - srcProfitAfterTaxTy],
    ['D44', d44 - srcCitTy], ['D45', d45 - srcVatPayTy]
  ];
  const badComponents = componentDiffs.filter(x => Math.abs(x[1]) > tolTy);
  let rootCause;

  if (badComponents.length) {
    rootCause = 'Sai liên kết tại: ' + badComponents.map(x => `${x[0]} lệch ${FS99Q_fmtTy_(x[1])}`).join('; ') + '.';
    results.push(['TỔNG HỢP - KẾT LUẬN', 'Đối chiếu chỉ tiêu Sheet 00', d25, summaryRight, summaryDiff,
      'FAIL', 'Có chỉ tiêu Sheet 00 không khớp nguồn chi tiết.', rootCause]);
  } else if (Math.abs(summaryDiff) > tolTy) {
    rootCause = 'Các chỉ tiêu nguồn đều khớp. Chênh lệch mục 14 là chênh lệch phạm vi kinh tế, không ép bằng 0.';
    results.push(['TỔNG HỢP - KẾT LUẬN', 'Đối chiếu phạm vi mục 14', d25, summaryRight, summaryDiff,
      'INFO', 'Chênh lệch chỉ để tham chiếu.', rootCause]);
  } else {
    rootCause = 'Các chỉ tiêu khớp nguồn và không có chênh lệch.';
    results.push(['TỔNG HỢP - KẾT LUẬN', 'Đối chiếu phạm vi mục 14', d25, summaryRight, summaryDiff,
      'PASS', 'Đơn vị tỷ đồng.', rootCause]);
  }

  const startRow = 50;
  const cols = 8;
  shC.getRange(startRow, 1, Math.max(1, shC.getMaxRows() - startRow + 1), cols).clearContent().clearFormat();
  shC.getRange(startRow, 1).setValue('AUDIT RECONCILIATION - XÁC ĐỊNH ĐÚNG VỊ TRÍ CẦN SỬA')
    .setFontWeight('bold').setFontSize(14);
  shC.getRange(startRow + 1, 1, 1, cols)
    .setValues([['Nhóm','Phép kiểm tra','Nguồn A','Nguồn B','Chênh lệch','Trạng thái','Ghi chú','Kết luận / vị trí cần sửa']])
    .setFontWeight('bold').setBackground('#d9ead3');
  shC.getRange(startRow + 2, 1, results.length, cols).setValues(results);
  shC.getRange(startRow + 2, 3, results.length, 3).setNumberFormat('#,##0.000');
  shC.getRange(startRow + 2, 6, results.length, 1).setFontWeight('bold');
  shC.getRange(startRow + 2, 8, results.length, 1).setWrap(true);
  shC.autoResizeColumns(1, 7);
  shC.setColumnWidth(8, 480);

  const fails = results.filter(r => r[5] === 'FAIL');
  const infos = results.filter(r => r[5] === 'INFO');
  const msg = fails.length
    ? 'AUDIT FAILED: ' + fails.length + ' sai lệch. Xem cột H tại Sheet 99. Checks.'
    : infos.length
      ? 'AUDIT PASS CÁC NGUỒN CHI TIẾT. Mục 14 chỉ còn chênh lệch phạm vi.'
      : 'AUDIT PASS: các phép đối chiếu chính đều khớp.';
  SpreadsheetApp.getUi().alert(msg);
  return { pass: fails.length === 0, failures: fails, information: infos, rootCause: rootCause };
}

function FS99Q_detectHeader_(sh, requiredNames) {
  const maxRows = Math.min(5, sh.getLastRow());
  let best = null;
  for (let row = 1; row <= maxRows; row++) {
    const headers = FS99Q_headers_(sh, row);
    const score = requiredNames.reduce((s, name) => s + (FS99Q_col_(headers, name) > 0 ? 1 : 0), 0);
    if (!best || score > best.score) best = { headerRow: row, headers: headers, score: score };
  }
  if (!best || best.score === 0) {
    throw new Error('Không nhận diện được dòng tiêu đề tại sheet ' + sh.getName() + '.');
  }
  return best;
}

function FS99Q_headers_(sh, row) {
  return sh.getRange(row, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(FS99Q_key_);
}
function FS99Q_col_(headers, name) {
  const i = headers.indexOf(FS99Q_key_(name));
  return i < 0 ? 0 : i + 1;
}
function FS99Q_sumCol_(sh, startRow, col) {
  if (!col || sh.getLastRow() < startRow) return 0;
  return sh.getRange(startRow, col, sh.getLastRow() - startRow + 1, 1).getValues()
    .reduce((s, r) => s + FS99Q_num_(r[0]), 0);
}
function FS99Q_num_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  const s = String(v == null ? '' : v).trim().replace(/\s/g, '').replace(/\./g, '').replace(',', '.');
  const n = Number(s);
  return isFinite(n) ? n : 0;
}
function FS99Q_key_(v) {
  return String(v || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd').replace(/[^a-z0-9]+/g, ' ').trim();
}
function FS99Q_fmtTy_(v) {
  return FS99Q_num_(v).toLocaleString('vi-VN', { minimumFractionDigits: 3, maximumFractionDigits: 3 }) + ' tỷ đồng';
}
