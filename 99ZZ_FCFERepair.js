/*************************************************
 * 99ZZ_FCFERepair.js
 * Sửa trực tiếp công thức FCFE trên Sheet 04 hiện tại.
 *************************************************/

function FS99ZZ_suaFCFE_HienTai() {
  const ss = SpreadsheetApp.getActive();
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');
  const tech = ss.getSheetByName('01. Kỹ thuật');

  if (!sh04 || !tech) {
    throw new Error('Thiếu Sheet 04 hoặc Sheet 01.');
  }

  const soThang = Number(FS04_getInfoValue_(tech, 'Số tháng mô hình')) || 0;
  if (!soThang) throw new Error('Thiếu Số tháng mô hình.');

  // AK = P + V - X = FCFF + giải ngân vay - trả gốc.
  sh04.getRange(3, 37, soThang, 1)
    .setFormulaR1C1('=RC[-21]+RC[-15]-RC[-13]');

  // AM = lũy kế FCFE.
  sh04.getRange(3, 39, soThang, 1)
    .setFormulaR1C1('=SUM(R3C37:RC37)');

  sh04.getRange(2, 37).setValue('FCFE khả dụng cho CSH');
  sh04.getRange(2, 39).setValue('FCFE lũy kế');

  SpreadsheetApp.flush();
  SpreadsheetApp.getActive().toast('Đã sửa FCFE: AK = P + V - X.', 'FS', 8);
}
