/*************************************************
 * 03C_CostOnlyEntry.gs
 * Entry point lập Sheet 03 theo nguyên tắc:
 * - Sheet 03 chỉ xác định chi phí, VAT và dòng tiền trước tài trợ.
 * - Toàn bộ trạng thái tài trợ do Sheet 04 tính rồi đồng bộ ngược.
 *************************************************/

function FS_lapSheet03_CostOnly() {
  FS_lapSheet03_Patched();

  const ss = SpreadsheetApp.getActive();
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  if (!sh03 || sh03.getLastRow() < 2) return;

  const numRows = sh03.getLastRow() - 1;

  // U:AE là khối tài trợ. Không để lại bộ số liệu sơ bộ do logic cũ của Sheet 03 tạo ra.
  // Khối này sẽ được điền bởi FS03_capNhatNguonVonTuSheet04_V2() sau khi Sheet 04 chạy.
  sh03.getRange(2, 21, numRows, 11).clearContent();

  // Khởi tạo số 0 để Sheet 02 có thể đọc chi phí lãi vay ở vòng chạy đầu tiên.
  const zeros = Array.from({ length: numRows }, () => Array(11).fill(0));
  sh03.getRange(2, 21, numRows, 11).setValues(zeros);

  // Ghi chú rõ nguồn dữ liệu tài trợ, không thay đổi cấu trúc sheet.
  const noteCol = 32; // AF - Ghi chú 1
  const notes = Array.from({ length: numRows }, () => ['Khối tài trợ do Sheet 04 tính và đồng bộ sau vòng hội tụ.']);
  sh03.getRange(2, noteCol, numRows, 1).setValues(notes);

  SpreadsheetApp.flush();
}
