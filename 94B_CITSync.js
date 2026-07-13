/*************************************************
 * 94B_CITSync.js
 * Đồng bộ Thuế TNDN theo tháng từ Sheet 02 sang Sheet 03 và Sheet 04
 * trước khi chạy chốt kiểm tra và lập báo cáo tổng hợp.
 *************************************************/

function FS94_syncCITFromSheet02_() {
  const ss = SpreadsheetApp.getActive();
  const sh02 = ss.getSheetByName('02. Doanh thu');
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');

  if (!sh02 || !sh03 || !sh04) {
    throw new Error('Thiếu Sheet 02, 03 hoặc 04 để đồng bộ Thuế TNDN.');
  }

  const taxByMonth = {};
  const lastRow02 = sh02.getLastRow();

  if (lastRow02 >= 2) {
    const rows02 = sh02.getRange(2, 1, lastRow02 - 1, 32).getValues();
    rows02.forEach(r => {
      const month = Number(r[0]) || 0;   // A - Tháng số
      const cit = Number(r[31]) || 0;    // AF - Thuế TNDN tạm tính
      if (!month) return;
      taxByMonth[month] = (taxByMonth[month] || 0) + Math.max(0, cit);
    });
  }

  FS94_syncCITToSheet_(sh03, 2, 1, 7, taxByMonth);   // G
  FS94_syncCITToSheet_(sh04, 3, 1, 12, taxByMonth);  // L
  SpreadsheetApp.flush();
}

function FS94_syncCITToSheet_(sheet, startRow, monthCol, taxCol, taxByMonth) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;

  const rowCount = lastRow - startRow + 1;
  const width = Math.max(monthCol, taxCol);
  const values = sheet.getRange(startRow, 1, rowCount, width).getValues();
  const out = values.map(r => {
    const month = Number(r[monthCol - 1]) || 0;
    return [month ? (taxByMonth[month] || 0) : 0];
  });

  sheet.getRange(startRow, taxCol, rowCount, 1).setValues(out);
}

function FS_chayMoHinh_Buoc3_SafeCIT() {
  const state = FS_getRunState_();
  if (state.status !== 'CONVERGED') {
    return FS_chayMoHinh_Buoc3();
  }

  FS94_syncCITFromSheet02_();
  return FS_chayMoHinh_Buoc3();
}
