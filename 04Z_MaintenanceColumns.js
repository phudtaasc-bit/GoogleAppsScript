/*************************************************
 * 04Z_MaintenanceColumns.js
 * Bổ sung các cột chi phí vận hành và bảo trì vào Sheet 04
 * mà không thay đổi các cột lõi A:AM đang dùng cho NPV/IRR.
 *
 * AN: Chi phí vận hành thuê trước VAT
 * AO: Chi phí bảo trì trước VAT
 * AP: Tổng chi phí vận hành & bảo trì trước VAT
 *************************************************/

const FS04M_BASE_LAP_SHEET04_ = FS_lapSheet04;

FS_lapSheet04 = function() {
  FS04M_BASE_LAP_SHEET04_();
  FS04M_applyMaintenanceColumns_();
};

function FS04M_applyMaintenanceColumns_() {
  const ss = SpreadsheetApp.getActive();
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');
  if (!sh03 || !sh04 || sh04.getLastRow() < 3) return;

  if (sh04.getMaxColumns() < 42) {
    sh04.insertColumnsAfter(sh04.getMaxColumns(), 42 - sh04.getMaxColumns());
  }

  const endRow = sh04.getLastRow();
  const rowCount = endRow - 2;
  const months = sh04.getRange(3, 1, rowCount, 1).getValues();

  const map03 = {};
  if (sh03.getLastRow() >= 2 && sh03.getLastColumn() >= 38) {
    const data03 = sh03.getRange(2, 1, sh03.getLastRow() - 1, 38).getValues();
    data03.forEach(r => {
      const t = Number(r[0]) || 0;
      if (!t) return;
      map03[t] = {
        operating: Number(r[35]) || 0,
        maintenance: Number(r[36]) || 0,
        combined: Number(r[37]) || 0
      };
    });
  }

  const outOperating = [];
  const outMaintenance = [];
  const outCombined = [];

  months.forEach(r => {
    const t = Number(r[0]) || 0;
    const item = map03[t] || { operating: 0, maintenance: 0, combined: 0 };
    outOperating.push([item.operating]);
    outMaintenance.push([item.maintenance]);
    outCombined.push([item.combined]);
  });

  sh04.getRange(1, 40, 1, 3).merge().setValue('CHI PHÍ VẬN HÀNH & BẢO TRÌ');
  sh04.getRange(2, 40, 1, 3).setValues([[
    'Chi phí vận hành thuê trước VAT',
    'Chi phí bảo trì trước VAT',
    'Tổng chi phí vận hành & bảo trì trước VAT'
  ]]);

  sh04.getRange(3, 40, rowCount, 1).setValues(outOperating);
  sh04.getRange(3, 41, rowCount, 1).setValues(outMaintenance);
  sh04.getRange(3, 42, rowCount, 1).setValues(outCombined);

  sh04.getRange(1, 40, 2, 3)
    .setFontWeight('bold')
    .setBackground('#fce5cd')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);

  sh04.getRange(3, 40, rowCount, 3)
    .setFontFamily('Arial')
    .setFontSize(10)
    .setNumberFormat('#,##0')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);

  sh04.autoResizeColumns(40, 3);
}
