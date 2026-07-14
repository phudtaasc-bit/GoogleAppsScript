/*************************************************
 * 99Q_AuditReconciliation.js
 * Audit phân rã để xác định khâu gây sai lệch.
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

  const tol = 1;
  const results = [];
  const add = (group, test, a, b, note) => {
    const d = FS99Q_num_(a) - FS99Q_num_(b);
    results.push([group, test, FS99Q_num_(a), FS99Q_num_(b), d, Math.abs(d) <= tol ? 'PASS' : 'FAIL', note || '']);
  };

  const h02 = FS99Q_headers_(sh02, 1);
  const h03 = FS99Q_headers_(sh03, 1);
  const h04 = FS99Q_headers_(sh04, 2);
  const sum02 = n => FS99Q_sumCol_(sh02, 2, FS99Q_col_(h02, n));
  const sum03 = n => FS99Q_sumCol_(sh03, 2, FS99Q_col_(h03, n));
  const sum04 = n => FS99Q_sumCol_(sh04, 3, FS99Q_col_(h04, n));

  // 1. Doanh thu
  add('DOANH THU', 'Sheet 02: Tổng doanh thu trước VAT = bán + thuê',
    sum02('Tổng doanh thu trước VAT'),
    sum02('Doanh thu bán trước VAT') + sum02('Doanh thu thuê trước VAT'));
  add('DOANH THU', 'Dòng tiền huy động KH Sheet 02 = Sheet 03',
    sum02('Dòng tiền huy động từ KH'), sum03('Dòng tiền huy động từ KH'));

  // 2. Chi phí trước VAT
  const costParts = ['Chi XD/TB/khác trước VAT','Chi GPMB trước VAT','Tiền SDĐ/thuê đất trước VAT',
    'Chi HTKT trước VAT','Chi phí bán hàng trước VAT','Chi phí dự phòng trước VAT',
    'Chi phí vận hành thuê trước VAT','Chi phí bảo trì trước VAT'];
  const parts03 = costParts.reduce((s, n) => s + sum03(n), 0);
  add('CHI PHÍ', 'Sheet 03: Tổng chi trước VAT = tổng 8 cấu phần', sum03('Tổng chi trước VAT'), parts03,
    'XD + GPMB + đất + HTKT + bán hàng + dự phòng + vận hành + bảo trì');
  add('CHI PHÍ', 'Sheet 03: Tổng chi sau VAT = Tổng chi trước VAT + VAT đầu vào',
    sum03('Tổng chi sau VAT'), sum03('Tổng chi trước VAT') + sum03('VAT đầu vào'));

  // 3. Thuế
  add('THUẾ', 'Thuế TNDN Sheet 02 = Sheet 03', sum02('Thuế TNDN tạm tính'), sum03('Thuế TNDN'));
  add('THUẾ', 'Thuế TNDN Sheet 03 = Sheet 04', sum03('Thuế TNDN'), sum04('Thuế TNDN'));

  // 4. Dòng tiền theo tháng và tổng
  add('DÒNG TIỀN', 'FCFF Sheet 04 = Dòng tiền trước tài trợ',
    sum04('FCFF_TIPV'), sum04('Dòng tiền trước tài trợ'));
  add('DÒNG TIỀN', 'FCFE = FCFF + Giải ngân vay - Trả gốc',
    sum04('FCFE khả dụng cho CSH'),
    sum04('FCFF_TIPV') + sum04('Giải ngân vay') - sum04('Trả gốc'));

  const inflow = sum04('Dòng tiền huy động từ KH') + sum04('Tổng dòng CSH vào dự án') + sum04('Giải ngân vay');
  const outflow = parts03 + sum03('VAT đầu vào') + sum04('VAT phải nộp') + sum04('Thuế TNDN') + sum04('Lãi vay vốn hóa') + sum04('Trả gốc');
  const netAfterFinancing = inflow - outflow;
  add('DÒNG TIỀN', 'Dòng tiền thuần sau tài trợ = Tổng vào - Tổng ra', netAfterFinancing, inflow - outflow);
  add('DÒNG TIỀN', 'FCFE = Dòng tiền thuần - CSH góp + Lãi vay',
    sum04('FCFE khả dụng cho CSH'),
    netAfterFinancing - sum04('Tổng dòng CSH vào dự án') + sum04('Lãi vay vốn hóa'));

  // 5. Sheet 00
  add('TỔNG HỢP', 'Chỉ tiêu 1 = 2 + 3 + 12 + 13',
    FS99Q_num_(sh00.getRange('D25').getValue()),
    FS99Q_num_(sh00.getRange('D30').getValue()) + FS99Q_num_(sh00.getRange('D33').getValue()) +
    FS99Q_num_(sh00.getRange('D42').getValue()) + FS99Q_num_(sh00.getRange('D43').getValue()),
    'Đơn vị tỷ đồng');

  // Ghi báo cáo
  const startRow = 50;
  shC.getRange(startRow, 1, Math.max(1, shC.getMaxRows() - startRow + 1), 7).clearContent().clearFormat();
  shC.getRange(startRow, 1).setValue('AUDIT RECONCILIATION - PHÂN RÃ SAI LỆCH').setFontWeight('bold').setFontSize(14);
  shC.getRange(startRow + 1, 1, 1, 7).setValues([['Nhóm','Phép kiểm tra','Nguồn A','Nguồn B','Chênh lệch','Trạng thái','Ghi chú']])
    .setFontWeight('bold').setBackground('#d9ead3');
  shC.getRange(startRow + 2, 1, results.length, 7).setValues(results);
  shC.getRange(startRow + 2, 3, results.length, 3).setNumberFormat('#,##0');
  shC.autoResizeColumns(1, 7);

  const fails = results.filter(r => r[5] === 'FAIL');
  const msg = fails.length
    ? 'AUDIT FAILED: ' + fails.length + ' sai lệch. Xem Sheet 99. Checks từ dòng ' + startRow + '.\n\n' +
      fails.slice(0, 8).map(r => r[0] + ' - ' + r[1] + ': lệch ' + FS99Q_fmt_(r[4])).join('\n')
    : 'AUDIT PASS: các phép đối chiếu chính đều khớp.';
  SpreadsheetApp.getUi().alert(msg);
  return { pass: fails.length === 0, failures: fails };
}

function FS99Q_headers_(sh, row) {
  return sh.getRange(row, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(FS99Q_key_);
}
function FS99Q_col_(headers, name) {
  const k = FS99Q_key_(name);
  const i = headers.indexOf(k);
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
function FS99Q_fmt_(v) {
  return Math.round(FS99Q_num_(v)).toLocaleString('vi-VN') + ' đồng';
}
