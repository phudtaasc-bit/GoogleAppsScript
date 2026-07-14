/*************************************************
 * ZZZZZZZZ_SummaryRevenueCheckFix.js
 *
 * Nguyên tắc:
 * - Chỉ tiêu 1 trên Sheet 00 vẫn lấy trực tiếp từ Sheet 02.
 * - Đẳng thức 1 = 2 + 3 + 12 + 13 chỉ là chỉ tiêu kiểm tra riêng,
 *   không được dùng để ghi đè doanh thu.
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
    'Đã khôi phục doanh thu liên kết từ Sheet 02 và thêm dòng kiểm tra cân đối riêng.'
  );
}

function FSZZZZZZZZ_restoreSummaryRevenueAndCheck_() {
  const ss = SpreadsheetApp.getActive();
  const sh00 = ss.getSheetByName('00. Tổng hợp');
  const sh02 = ss.getSheetByName('02. Doanh thu');
  if (!sh00 || !sh02) {
    throw new Error('Thiếu Sheet 00. Tổng hợp hoặc Sheet 02. Doanh thu.');
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

  // Dòng kiểm tra riêng, không tham gia tạo ra số liệu của chỉ tiêu 1.
  const checkRow = 44;
  sh00.getRange(checkRow, 1, 1, 5).clearContent();
  sh00.getRange(checkRow, 1).setValue('14');
  sh00.getRange(checkRow, 2).setValue('Chênh lệch kiểm tra cân đối');
  sh00.getRange(checkRow, 3).setValue('tỷ đồng');
  sh00.getRange(checkRow, 4).setFormula('=D25-(D30+D33+D42+D43)');
  sh00.getRange(checkRow, 5).setValue('Kiểm tra: Chỉ tiêu 1 - (2 + 3 + 12 + 13); kết quả chuẩn = 0');

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
  sh00.setRowHeight(checkRow, 30);
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
