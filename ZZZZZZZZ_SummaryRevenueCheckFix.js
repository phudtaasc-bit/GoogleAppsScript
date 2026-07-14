/*************************************************
 * ZZZZZZZZ_SummaryRevenueCheckFix.js
 *
 * Nguyên tắc:
 * - Chỉ tiêu 1 trên Sheet 00 vẫn lấy trực tiếp từ Sheet 02.
 * - Đẳng thức 1 = 2 + 3 + 12 + 13 chỉ là phép đối chiếu khác phạm vi,
 *   không được dùng để ghi đè doanh thu hoặc ép kết quả bằng 0.
 * - Bổ sung trình bày riêng chi phí vận hành và chi phí bảo trì,
 *   không chèn dòng làm thay đổi địa chỉ các chỉ tiêu hiện hữu.
 *************************************************/

const FSZZZZZZZZ_BASE_LAP_SHEET00_ = FS_lapSheet00;

FS_lapSheet00 = function() {
  FSZZZZZZZZ_BASE_LAP_SHEET00_();
  FSZZZZZZZZ_restoreSummaryRevenueAndCheck_();
};

/**
 * Chạy riêng khi chỉ cần sửa Sheet 00 hiện tại, không lập lại mô hình.
 */
function FSZZZZZZZZ_suaTongHopDoanhThuVaKiemTra() {
  FSZZZZZZZZ_restoreSummaryRevenueAndCheck_();
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    'Đã cập nhật doanh thu liên kết, tiêu chí đối chiếu và chi phí vận hành/bảo trì trên Sheet 00.'
  );
}

function FSZZZZZZZZ_restoreSummaryRevenueAndCheck_() {
  const ss = SpreadsheetApp.getActive();
  const sh00 = ss.getSheetByName('00. Tổng hợp');
  const sh02 = ss.getSheetByName('02. Doanh thu');
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  if (!sh00 || !sh02 || !sh03) {
    throw new Error('Thiếu Sheet 00. Tổng hợp, Sheet 02. Doanh thu hoặc Sheet 03. Chi phí & Vốn.');
  }

  const headers02 = sh02
    .getRange(1, 1, 1, sh02.getLastColumn())
    .getDisplayValues()[0];

  const cashCol = FSZZZZZZZZ_findHeader_(headers02, [
    'Dòng tiền huy động từ KH',
    'Dòng tiền huy động từ khách hàng'
  ]);

  if (cashCol < 1) {
    throw new Error('Không tìm thấy cột Dòng tiền huy động từ KH trên Sheet 02.');
  }

  const cashLetter = FSZZZZZZZZ_colLetter_(cashCol);
  const s02 = `'02. Doanh thu'`;

  // Chỉ tiêu 1: luôn link trực tiếp từ Sheet chi tiết.
  sh00.getRange('D25').setFormula(
    `=SUM(${s02}!${cashLetter}2:${cashLetter})/1000000000`
  );

  // D44 chỉ phản ánh chênh lệch giữa các chỉ tiêu không đồng nhất phạm vi kinh tế.
  const checkRow = 44;
  sh00.getRange(checkRow, 1, 1, 5).clearContent();
  sh00.getRange(checkRow, 1).setValue('14');
  sh00.getRange(checkRow, 2).setValue('Chênh lệch đối chiếu phạm vi');
  sh00.getRange(checkRow, 3).setValue('tỷ đồng');
  sh00.getRange(checkRow, 4).setFormula('=D25-(D30+D33+D42+D43)');
  sh00.getRange(checkRow, 5).setValue(
    'Tham chiếu: Doanh thu có VAT - (Tổng chi phí có VAT + LNST + Thuế TNDN + VAT phải nộp). ' +
    'Các chỉ tiêu không cùng phạm vi kinh tế; kết quả không bắt buộc bằng 0 và không dùng để điều chỉnh số liệu.'
  );

  sh00.getRange(checkRow, 1, 1, 5)
    .setFontFamily('Times New Roman')
    .setFontSize(11)
    .setBorder(true, true, true, true, true, true)
    .setVerticalAlignment('middle');

  sh00.getRange(checkRow, 1, 1, 3).setFontWeight('bold');
  sh00.getRange(checkRow, 4)
    .setNumberFormat('#,##0.0;[Red]-#,##0.0')
    .setFontWeight('bold');
  sh00.getRange(checkRow, 5).setWrap(true);
  sh00.setRowHeight(checkRow, 48);

  // Bổ sung hai cấu phần chi phí sau bảng hiện hữu để không làm dịch chuyển D25:D44.
  const headers03 = sh03
    .getRange(1, 1, 1, sh03.getLastColumn())
    .getDisplayValues()[0];
  const opCol = FSZZZZZZZZ_findHeader_(headers03, [
    'Chi phí vận hành thuê trước VAT',
    'Chi phí vận hành trước VAT'
  ]);
  const maintCol = FSZZZZZZZZ_findHeader_(headers03, [
    'Chi phí bảo trì trước VAT'
  ]);

  const detailRows = [
    [45, '2.3', 'Chi phí vận hành (trước VAT)', opCol],
    [46, '2.4', 'Chi phí bảo trì (trước VAT)', maintCol]
  ];

  detailRows.forEach(item => {
    const row = item[0];
    const tt = item[1];
    const label = item[2];
    const col = item[3];

    sh00.getRange(row, 1, 1, 5).clearContent();
    sh00.getRange(row, 1).setValue(tt);
    sh00.getRange(row, 2).setValue(label);
    sh00.getRange(row, 3).setValue('tỷ đồng');

    if (col > 0) {
      const letter = FSZZZZZZZZ_colLetter_(col);
      sh00.getRange(row, 4).setFormula(
        `=SUM('03. Chi phí & Vốn'!${letter}2:${letter})/1000000000`
      );
    } else {
      sh00.getRange(row, 4).setValue(0);
    }

    sh00.getRange(row, 5).setValue('Nguồn: Sheet 03. Chi phí & Vốn');
    sh00.getRange(row, 1, 1, 5)
      .setFontFamily('Times New Roman')
      .setFontSize(11)
      .setBorder(true, true, true, true, true, true)
      .setVerticalAlignment('middle');
    sh00.getRange(row, 4).setNumberFormat('#,##0.0;[Red]-#,##0.0');
  });
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
