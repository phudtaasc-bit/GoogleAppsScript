/*************************************************
 * 03_04_Format.gs
 * Định dạng Sheet 03 - không sửa công thức
 *************************************************/

function FS03_formatV2() {
  const ss = SpreadsheetApp.getActive();
  const sh = FS_getSheetByPrefix_(ss, '03');
  if (!sh) throw new Error('Không tìm thấy sheet bắt đầu bằng "03".');

  const lastRow = sh.getLastRow();
  const lastCol = 35; // A:AI

  if (lastRow < 2) throw new Error('Sheet 03 chưa có dữ liệu.');

  sh.showColumns(1, sh.getMaxColumns());
  sh.showRows(1, sh.getMaxRows());

  // Nếu sheet đang có 1 hàng tiêu đề, chèn thêm hàng group ở trên.
  const a1 = String(sh.getRange(1, 1).getDisplayValue() || '').trim();
  if (a1 === 'Tháng số') {
    sh.insertRowBefore(1);
  }

  const newLastRow = sh.getLastRow();

  // Clear format, merge cũ ở hàng 1
  sh.getRange(1, 1, newLastRow, lastCol).breakApart();
  sh.getRange(1, 1, newLastRow, lastCol).clearFormat();

  FS_UI.base(sh, newLastRow, lastCol);

  // Hàng 1 - nhóm cột
  FS_UI.titleRow(sh, 1, 1, 4,  'THỜI GIAN', FS_UI.COLOR.TIME);
  FS_UI.titleRow(sh, 1, 5, 3,  'DÒNG TIỀN HOẠT ĐỘNG', FS_UI.COLOR.OPERATING);
  FS_UI.titleRow(sh, 1, 8, 9,  'CHI PHÍ DỰ ÁN', FS_UI.COLOR.COST);
  FS_UI.titleRow(sh, 1, 17, 3, 'THUẾ GTGT', FS_UI.COLOR.VAT);
  FS_UI.titleRow(sh, 1, 20, 1, 'DÒNG TIỀN TRƯỚC TÀI TRỢ', FS_UI.COLOR.KPI);
  FS_UI.titleRow(sh, 1, 21, 5, 'NGUỒN VỐN', FS_UI.COLOR.FUNDING);
  FS_UI.titleRow(sh, 1, 26, 4, 'KHOẢN VAY', FS_UI.COLOR.DEBT);
  FS_UI.titleRow(sh, 1, 30, 2, 'KẾT QUẢ DÒNG TIỀN', FS_UI.COLOR.CASH);
  FS_UI.titleRow(sh, 1, 32, 4, 'GHI CHÚ', '#F2F2F2');

  // Hàng 2 - header từng cột
  FS_UI.header(sh, 2, 1, 4, FS_UI.COLOR.TIME);
  FS_UI.header(sh, 2, 5, 3, FS_UI.COLOR.OPERATING);
  FS_UI.header(sh, 2, 8, 9, FS_UI.COLOR.COST);
  FS_UI.header(sh, 2, 17, 3, FS_UI.COLOR.VAT);
  FS_UI.header(sh, 2, 20, 1, FS_UI.COLOR.KPI);
  FS_UI.header(sh, 2, 21, 5, FS_UI.COLOR.FUNDING);
  FS_UI.header(sh, 2, 26, 4, FS_UI.COLOR.DEBT);
  FS_UI.header(sh, 2, 30, 2, FS_UI.COLOR.CASH);
  FS_UI.header(sh, 2, 32, 4, '#F2F2F2');

  const dataRows = Math.max(0, newLastRow - 2);

  // Format dữ liệu
  if (dataRows > 0) {
    FS_UI.block(sh, 3, 1, dataRows, 4, '#FFFFFF');
    FS_UI.block(sh, 3, 5, dataRows, 3, '#FFFFFF');
    FS_UI.block(sh, 3, 8, dataRows, 9, '#FFFFFF');
    FS_UI.block(sh, 3, 17, dataRows, 3, '#FFFFFF');
    FS_UI.block(sh, 3, 20, dataRows, 1, '#FFFFFF');
    FS_UI.block(sh, 3, 21, dataRows, 5, '#FFFFFF');
    FS_UI.block(sh, 3, 26, dataRows, 4, '#FFFFFF');
    FS_UI.block(sh, 3, 30, dataRows, 2, '#FFFFFF');

    // Định dạng số
    FS_UI.numberFormat(sh, 3, 1, dataRows, 1, '0');
    FS_UI.numberFormat(sh, 3, 2, dataRows, 1, 'dd/mm/yyyy');
    FS_UI.numberFormat(sh, 3, 3, dataRows, 1, '0');
    FS_UI.numberFormat(sh, 3, 5, dataRows, 31, '#,##0');

    // Căn giữa thời gian
    sh.getRange(3, 1, dataRows, 4).setHorizontalAlignment('center');

    // KPI columns: T, U, AD, AE
    FS_UI.kpiColumn(sh, 2, 20, dataRows + 1, '#B7DEE8'); // T - Dòng tiền trước tài trợ
    FS_UI.kpiColumn(sh, 2, 21, dataRows + 1, '#F8CBAD'); // U - Nhu cầu vốn
    FS_UI.kpiColumn(sh, 2, 30, dataRows + 1, '#C6E0B4'); // AD - Dòng tiền sau tài trợ
    FS_UI.kpiColumn(sh, 2, 31, dataRows + 1, '#A9D18E'); // AE - Tiền cuối kỳ
  }

  // Freeze
  FS_UI.freeze(sh, 2, 4);

  // Chiều cao header
  sh.setRowHeight(1, 28);
  sh.setRowHeight(2, 54);

  // Auto width
  FS_UI.autoWidth(sh, 1, lastCol);

  // Ẩn cột kỹ thuật mặc định theo chế độ quản trị
  FS03_viewQuanTri();

  SpreadsheetApp.getUi().alert('Đã định dạng Sheet 03 V2.');
}