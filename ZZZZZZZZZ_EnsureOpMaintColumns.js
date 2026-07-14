/*************************************************
 * ZZZZZZZZZ_EnsureOpMaintColumns.js
 * Bảo đảm Sheet 03 luôn có cột chi phí vận hành/bảo trì trước khi đối chiếu VAT.
 * Không thay đổi cấu trúc các cột hiện hữu; chỉ bổ sung ở cuối sheet khi thiếu.
 *************************************************/

const FS03_REBUILD_COSTS_LEGACY_ = FSZZZZZZZ_rebuildSheet03Costs_;

FSZZZZZZZ_rebuildSheet03Costs_ = function(sh03, vatRates, opVatRate) {
  FS03_ensureOperatingMaintenanceColumns_(sh03);
  return FS03_REBUILD_COSTS_LEGACY_(sh03, vatRates, opVatRate);
};

function FS03_ensureOperatingMaintenanceColumns_(sh03) {
  const ss = sh03.getParent();
  const sh02 = ss.getSheetByName('02. Doanh thu');
  if (!sh02) throw new Error('Không tìm thấy Sheet 02. Doanh thu để đồng bộ chi phí vận hành/bảo trì.');

  let headers03 = sh03.getRange(1, 1, 1, sh03.getLastColumn()).getDisplayValues()[0];
  let opCol03 = FS03_findHeader_(headers03, [
    'Chi phí vận hành thuê trước VAT',
    'Chi phí vận hành trước VAT'
  ]);
  let maintCol03 = FS03_findHeader_(headers03, [
    'Chi phí bảo trì trước VAT'
  ]);

  let nextCol = sh03.getLastColumn() + 1;
  if (opCol03 < 0) {
    sh03.getRange(1, nextCol).setValue('Chi phí vận hành thuê trước VAT');
    opCol03 = nextCol - 1;
    nextCol++;
  }
  if (maintCol03 < 0) {
    sh03.getRange(1, nextCol).setValue('Chi phí bảo trì trước VAT');
    maintCol03 = nextCol - 1;
    nextCol++;
  }

  const headers02 = sh02.getRange(1, 1, 1, sh02.getLastColumn()).getDisplayValues()[0];
  const monthCol02 = FS03_findHeader_(headers02, ['Tháng số']);
  const opCol02 = FS03_findHeader_(headers02, [
    'Chi phí vận hành thuê trước VAT',
    'Chi phí vận hành trước VAT'
  ]);
  const maintCol02 = FS03_findHeader_(headers02, ['Chi phí bảo trì trước VAT']);

  if (monthCol02 < 0 || opCol02 < 0 || maintCol02 < 0) {
    throw new Error('Sheet 02 chưa có đủ cột Tháng số/Chi phí vận hành/Chi phí bảo trì.');
  }

  const byMonth = {};
  const n02 = Math.max(0, sh02.getLastRow() - 1);
  if (n02 > 0) {
    const data02 = sh02.getRange(2, 1, n02, sh02.getLastColumn()).getValues();
    data02.forEach(r => {
      const month = Number(r[monthCol02]) || 0;
      if (!month) return;
      if (!byMonth[month]) byMonth[month] = { op: 0, maint: 0 };
      byMonth[month].op += Number(r[opCol02]) || 0;
      byMonth[month].maint += Number(r[maintCol02]) || 0;
    });
  }

  const n03 = Math.max(0, sh03.getLastRow() - 1);
  if (n03 > 0) {
    const months03 = sh03.getRange(2, 1, n03, 1).getValues();
    const opOut = [];
    const maintOut = [];
    months03.forEach(r => {
      const x = byMonth[Number(r[0]) || 0] || { op: 0, maint: 0 };
      opOut.push([x.op]);
      maintOut.push([x.maint]);
    });
    sh03.getRange(2, opCol03 + 1, n03, 1).setValues(opOut).setNumberFormat('#,##0');
    sh03.getRange(2, maintCol03 + 1, n03, 1).setValues(maintOut).setNumberFormat('#,##0');
  }

  sh03.getRange(1, opCol03 + 1, 1, 2)
    .setFontWeight('bold')
    .setWrap(true)
    .setBackground('#D9EAD3');
}

function FS03_findHeader_(headers, aliases) {
  const keys = aliases.map(FS03_headerKey_);
  return headers.findIndex(h => keys.includes(FS03_headerKey_(h)));
}

function FS03_headerKey_(v) {
  return String(v || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}
