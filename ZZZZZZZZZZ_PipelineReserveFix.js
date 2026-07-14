/*************************************************
 * ZZZZZZZZZZ_PipelineReserveFix.js
 * Vá trực tiếp Bước 8.1:
 * - Dự phòng tháng = tỷ lệ dự phòng × (Chi XD/TB + Chi HTKT).
 * - Tạo/đồng bộ cột CP vận hành và bảo trì trên Sheet 03.
 * - Tính lại tổng chi trước VAT, VAT đầu vào và tổng chi sau VAT.
 * Không lập Sheet 04 trong bước này để tránh timeout.
 *************************************************/

FS99P_buoc1_CapNhatChiPhiNguon = function() {
  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName('01. Kỹ thuật');
  const sh02 = ss.getSheetByName('02. Doanh thu');
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');

  if (!tech || !sh02 || !sh03) {
    throw new Error('Thiếu Sheet 01. Kỹ thuật, 02. Doanh thu hoặc 03. Chi phí & Vốn.');
  }

  FSZZZZZZZZZ_ensureCostColumns_(sh03);

  const reserveCfg = FSZZZZZZ_getReserveConfig_(tech);
  const reserveRate = Number(reserveCfg.rate) || 0;
  const opMaintByMonth = FSZZZZZZ_getOpMaintByMonth_(sh02);

  const headers = sh03.getRange(1, 1, 1, sh03.getLastColumn()).getDisplayValues()[0];
  const col = name => FSZZZZZZZZZ_headerIndex_(headers, name);

  const cMonth = col('Tháng số');
  const cXD = col('Chi XD/TB/khác trước VAT');
  const cHTKT = col('Chi HTKT trước VAT');
  const cReserve = col('Chi phí dự phòng trước VAT');
  const cOp = col('Chi phí vận hành thuê trước VAT');
  const cMaint = col('Chi phí bảo trì trước VAT');
  const cOpMaint = col('Tổng chi phí vận hành & bảo trì trước VAT');

  if ([cMonth, cXD, cHTKT, cReserve, cOp, cMaint, cOpMaint].some(v => v < 0)) {
    throw new Error('Sheet 03 thiếu cột nguồn để cập nhật dự phòng/vận hành/bảo trì.');
  }

  const lastRow = sh03.getLastRow();
  if (lastRow >= 2) {
    const n = lastRow - 1;
    const data = sh03.getRange(2, 1, n, sh03.getLastColumn()).getValues();
    const outReserve = [];
    const outOp = [];
    const outMaint = [];
    const outTotal = [];

    data.forEach(r => {
      const month = Number(r[cMonth]) || 0;
      const xd = Number(r[cXD]) || 0;
      const htkt = Number(r[cHTKT]) || 0;
      const x = opMaintByMonth[month] || { op: 0, maint: 0 };
      const op = Math.max(0, Number(x.op) || 0);
      const maint = Math.max(0, Number(x.maint) || 0);

      outReserve.push([Math.max(0, (xd + htkt) * reserveRate)]);
      outOp.push([op]);
      outMaint.push([maint]);
      outTotal.push([op + maint]);
    });

    sh03.getRange(2, cReserve + 1, n, 1).setValues(outReserve);
    sh03.getRange(2, cOp + 1, n, 1).setValues(outOp);
    sh03.getRange(2, cMaint + 1, n, 1).setValues(outMaint);
    sh03.getRange(2, cOpMaint + 1, n, 1).setValues(outTotal);
  }

  const vatRates = FSZZZZZZZ_readCommonCostVatRates_(tech);
  const opVatRate = FSZZZZZZZ_pickRate_(vatRates, [
    'Chi phí vận hành',
    'Chi phí bán hàng',
    'Chi phí XD/TB/khác'
  ]);

  FSZZZZZZZ_rebuildSheet02TaxCost_(sh02);
  FSZZZZZZZ_rebuildSheet03Costs_(sh03, vatRates, opVatRate);

  SpreadsheetApp.flush();
  ss.toast(
    'Bước 8.1 hoàn thành: đã cập nhật dự phòng theo tháng, CP vận hành/bảo trì, tổng chi và VAT đầu vào.',
    'FS Pipeline',
    8
  );
};

function FSZZZZZZZZZ_ensureCostColumns_(sh) {
  const required = [
    'Chi phí vận hành thuê trước VAT',
    'Chi phí bảo trì trước VAT',
    'Tổng chi phí vận hành & bảo trì trước VAT'
  ];

  let headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  required.forEach(name => {
    if (FSZZZZZZZZZ_headerIndex_(headers, name) >= 0) return;
    const newCol = sh.getLastColumn() + 1;
    sh.getRange(1, newCol).setValue(name);
    headers.push(name);
  });
}

function FSZZZZZZZZZ_headerIndex_(headers, name) {
  const key = FSZZZZZZZZZ_norm_(name);
  return headers.findIndex(h => FSZZZZZZZZZ_norm_(h) === key);
}

function FSZZZZZZZZZ_norm_(v) {
  return String(v || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/&/g, ' va ')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}
