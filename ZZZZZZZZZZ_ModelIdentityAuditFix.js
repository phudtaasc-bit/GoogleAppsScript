/*************************************************
 * ZZZZZZZZZZ_ModelIdentityAuditFix.js
 * - Không tạo dòng điều chỉnh để ép cân đối.
 * - Dùng dấu ; cho công thức theo locale Việt Nam.
 * - Mục 14 khác 0 => audit FAIL và phải kiểm tra lại toàn bộ mô hình.
 *************************************************/

const FS_IDENTITY_BASE_LAP_SHEET00_ = FS_lapSheet00;

FS_lapSheet00 = function() {
  FS_IDENTITY_BASE_LAP_SHEET00_();
  FS_IDENTITY_fixVietnameseFormulas_();
};

function FS_IDENTITY_fixVietnameseFormulas_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('00. Tổng hợp');
  if (!sh) throw new Error('Không tìm thấy Sheet 00. Tổng hợp.');
  for (let r = 6; r <= 11; r++) {
    sh.getRange(r, 4).setFormula(`=IFERROR(C${r}/$C$12;0)`);
  }
}

function FS99P_buoc5_KiemTraNhanh() {
  return FS99Q_chayAuditReconciliation();
}

function FS99Q_chayAuditReconciliation() {
  const ss = SpreadsheetApp.getActive();
  const sh02 = ss.getSheetByName('02. Doanh thu');
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');
  const sh04A = ss.getSheetByName('04A. TH dòng tiền');
  const sh00 = ss.getSheetByName('00. Tổng hợp');
  let shC = ss.getSheetByName('99. Checks');
  if (!sh02 || !sh03 || !sh04 || !sh00) {
    throw new Error('Thiếu Sheet 00, 02, 03 hoặc 04 để audit.');
  }
  if (!shC) shC = ss.insertSheet('99. Checks');

  SpreadsheetApp.flush();
  const TOL_DONG = 1;
  const TOL_TY = 0.001;
  const results = [];

  const add = (layer, test, a, b, note, fix, tol) => {
    const av = FS99Q_num_(a);
    const bv = FS99Q_num_(b);
    const diff = av - bv;
    const limit = tol == null ? TOL_DONG : tol;
    const status = Math.abs(diff) <= limit ? 'PASS' : 'FAIL';
    results.push([layer, test, av, bv, diff, status, note || '', status === 'FAIL' ? (fix || '') : '']);
  };

  const m02 = FS99Q_detectHeader_(sh02, ['Tháng số', 'Dòng tiền huy động từ KH', 'Thuế TNDN tạm tính']);
  const m03 = FS99Q_detectHeader_(sh03, ['Tháng số', 'Tổng chi sau VAT', 'Thuế TNDN']);
  const m04 = FS99Q_detectHeader_(sh04, ['Tháng số', 'FCFF_TIPV', 'Thuế TNDN']);
  const s02 = n => FS99Q_sumNamed_(sh02, m02, n);
  const s03 = n => FS99Q_sumNamed_(sh03, m03, n);
  const s04 = n => FS99Q_sumNamed_(sh04, m04, n);

  add('TẦNG 1 - DOANH THU', 'Tổng doanh thu trước VAT = doanh thu bán + doanh thu thuê',
    s02('Tổng doanh thu trước VAT'),
    s02('Doanh thu bán trước VAT') + s02('Doanh thu thuê trước VAT'),
    'Kiểm tra nội tại Sheet 02.', 'Sửa công thức doanh thu tại Sheet 02.', TOL_DONG);

  add('TẦNG 1 - LUỒNG DỮ LIỆU', 'Dòng tiền khách hàng Sheet 02 = Sheet 03',
    s02('Dòng tiền huy động từ KH'), s03('Dòng tiền huy động từ KH'),
    'Kiểm tra luồng 02 -> 03.', 'Sửa module cập nhật nguồn Sheet 03.', TOL_DONG);

  add('TẦNG 1 - CHI PHÍ', 'Tổng chi sau VAT Sheet 03 = Tổng chi trước VAT + VAT đầu vào',
    s03('Tổng chi sau VAT'), s03('Tổng chi trước VAT') + s03('VAT đầu vào'),
    'Kiểm tra nội tại Sheet 03.', 'Sửa Tổng chi sau VAT hoặc VAT đầu vào tại Sheet 03.', TOL_DONG);

  add('TẦNG 1 - THUẾ', 'Thuế TNDN Sheet 02 = Sheet 03',
    s02('Thuế TNDN tạm tính'), s03('Thuế TNDN'),
    'Kiểm tra luồng thuế 02 -> 03.', 'Sửa module đồng bộ Thuế TNDN.', TOL_DONG);

  add('TẦNG 1 - THUẾ', 'Thuế TNDN Sheet 03 = Sheet 04',
    s03('Thuế TNDN'), s04('Thuế TNDN'),
    'Kiểm tra luồng thuế 03 -> 04.', 'Sửa module lập Sheet 04.', TOL_DONG);

  add('TẦNG 1 - DÒNG TIỀN', 'FCFF = Dòng tiền trước tài trợ',
    s04('FCFF_TIPV'), s04('Dòng tiền trước tài trợ'),
    'Kiểm tra FCFF.', 'Sửa công thức FCFF Sheet 04.', TOL_DONG);

  add('TẦNG 1 - DÒNG TIỀN', 'FCFE = FCFF + giải ngân vay - trả gốc',
    s04('FCFE khả dụng cho CSH'), s04('FCFF_TIPV') + s04('Giải ngân vay') - s04('Trả gốc'),
    'Kiểm tra FCFE.', 'Sửa công thức FCFE Sheet 04.', TOL_DONG);

  if (sh04A) {
    const annual = FS99Q_readAnnualTotals_(sh04A);
    const totalIn = FS99Q_annualValue_(annual, 'TỔNG DÒNG TIỀN VÀO');
    const totalOut = FS99Q_annualValue_(annual, 'TỔNG DÒNG TIỀN RA');
    const net = FS99Q_annualValue_(annual, 'Dòng tiền thuần sau tài trợ');
    add('TẦNG 2 - SHEET 04A', 'Dòng tiền thuần sau tài trợ = Tổng vào - Tổng ra',
      net, totalIn - totalOut, 'Kiểm tra cầu nối Sheet 04A.', 'Sửa Sheet 04A.', TOL_TY);
  }

  const summary = FS99Q_readSummaryByLabel_(sh00);
  const revenue = FS99Q_summaryValue_(summary, 'Tổng doanh thu có VAT');
  const cost = FS99Q_summaryValue_(summary, 'Tổng chi phí có VAT');
  const profit = FS99Q_summaryValue_(summary, 'Lợi nhuận sau thuế');
  const cit = FS99Q_summaryValue_(summary, 'Tổng Thuế TNDN');
  const vat = FS99Q_summaryValue_(summary, 'Tổng VAT phải nộp');
  const check = FS99Q_summaryValue_(summary, 'Chênh lệch đối chiếu phạm vi');

  add('TẦNG 3 - SHEET 00', 'Tổng doanh thu có VAT = nguồn Sheet 02',
    revenue, s02('Dòng tiền huy động từ KH') / 1e9,
    'Đơn vị tỷ đồng.', 'Sửa chỉ tiêu Tổng doanh thu Sheet 00.', TOL_TY);

  add('TẦNG 3 - SHEET 00', 'Lợi nhuận sau thuế = nguồn Sheet 04',
    profit, s04('Lợi nhuận sau thuế') / 1e9,
    'Đơn vị tỷ đồng.', 'Sửa chỉ tiêu LNST Sheet 00.', TOL_TY);

  add('TẦNG 3 - SHEET 00', 'Thuế TNDN = nguồn Sheet 04',
    cit, s04('Thuế TNDN') / 1e9,
    'Đơn vị tỷ đồng.', 'Sửa chỉ tiêu Thuế TNDN Sheet 00.', TOL_TY);

  add('TẦNG 3 - SHEET 00', 'VAT phải nộp = nguồn Sheet 04',
    vat, s04('VAT phải nộp') / 1e9,
    'Đơn vị tỷ đồng.', 'Sửa chỉ tiêu VAT phải nộp Sheet 00.', TOL_TY);

  const identityRight = cost + profit + cit + vat;
  add('TẦNG 4 - CÂN ĐỐI TOÀN MÔ HÌNH',
    'Doanh thu có VAT = Chi phí có VAT + LNST + Thuế TNDN + VAT phải nộp',
    revenue, identityRight,
    'Đây là phép kiểm tra bắt buộc. Không tạo dòng điều chỉnh để ép cân đối.',
    'Dừng tại đây và phân rã lại toàn bộ doanh thu, chi phí, lợi nhuận, VAT và Thuế TNDN theo cùng phạm vi.', TOL_TY);

  add('TẦNG 4 - CÂN ĐỐI TOÀN MÔ HÌNH', 'Mục 14 phải bằng 0',
    check, 0,
    'Mục 14 là chênh lệch của phương trình cân đối toàn mô hình.',
    'Không sửa trực tiếp mục 14; sửa nguyên nhân tại các sheet nguồn.', TOL_TY);

  const failRows = results.filter(r => r[5] === 'FAIL');
  const rootCause = failRows.length
    ? 'Lỗi đầu tiên: ' + failRows[0][0] + ' - ' + failRows[0][1] + '. ' + failRows[0][7]
    : 'Toàn bộ nguồn chính và phương trình cân đối mô hình đều khớp; Mục 14 bằng 0.';

  results.push(['KẾT LUẬN', failRows.length ? 'Cần kiểm tra lại mô hình' : 'Kết quả audit',
    0, 0, 0, failRows.length ? 'FAIL' : 'PASS', '', rootCause]);

  FS99Q_writeResults_(shC, results);
  SpreadsheetApp.getUi().alert(failRows.length
    ? 'AUDIT FAILED.\n\n' + rootCause + '\n\nKhông được tạo dòng điều chỉnh để ép mục 14 về 0.'
    : 'AUDIT PASS: phương trình cân đối toàn mô hình khớp và Mục 14 bằng 0.');

  return { pass: failRows.length === 0, failures: failRows, rootCause };
}
