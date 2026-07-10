/*************************************************
 * 03C_ChiPhiCostOnly.gs
 * Lập Sheet 03 chỉ với chi phí, VAT và dòng tiền trước tài trợ.
 * Khối tài trợ U:AE luôn do Sheet 04 tính và đồng bộ ngược về Sheet 03.
 *************************************************/

function FS_lapSheet03_CostOnly() {
  FS_lapSheet03_Patched();

  const ss = SpreadsheetApp.getActive();
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  if (!sh03 || sh03.getLastRow() < 2) return;

  const numRows = sh03.getLastRow() - 1;

  // U:AE = 11 cột tài trợ. Không giữ kết quả sơ bộ từ logic cũ.
  sh03.getRange(2, 21, numRows, 11).clearContent();

  // Ghi chú rõ nguồn dữ liệu để tránh hiểu nhầm khi chạy riêng Sheet 03.
  sh03.getRange(2, 32, numRows, 1).setValues(
    Array.from({ length: numRows }, () => ['Tài trợ được tính tại Sheet 04'])
  );

  SpreadsheetApp.flush();
}
