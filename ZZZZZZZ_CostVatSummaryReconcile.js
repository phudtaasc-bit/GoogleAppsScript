/*************************************************
 * ZZZZZZZ_CostVatSummaryReconcile.js
 *
 * Mục tiêu:
 * 1) Tổng chi trước VAT Sheet 03 bao gồm đầy đủ CPVH và bảo trì.
 * 2) VAT đầu vào bao gồm VAT của CPVH và bảo trì theo tỷ lệ cấu hình.
 * 3) Giá vốn tính thuế Sheet 02 bao gồm chi phí bảo trì.
 * 4) Sheet 04 hiển thị đúng CPVH/bảo trì theo tháng.
 * 5) Sheet 00 bảo đảm chỉ tiêu 1 = 2 + 3 + 12 + 13.
 *************************************************/

function FSZZZZZZZ_reconcileCostsAndSummary() {
  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName('01. Kỹ thuật');
  const sh02 = ss.getSheetByName('02. Doanh thu');
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  if (!tech || !sh02 || !sh03) {
    throw new Error('Thiếu Sheet 01. Kỹ thuật, 02. Doanh thu hoặc 03. Chi phí & Vốn.');
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

  // Lập lại Sheet 04 trên dòng chi phí đã bao gồm dự phòng, CPVH, bảo trì và VAT.
  FS_lapSheet04();
  FSZZZZZZZ_syncOperatingMaintenanceTo04_(ss);

  // Bảo đảm FCFE không bị patch cũ ghi đè.
  FSZZZZZZZ_forceFCFE_(ss);

  FS_lapSheet04A();
  FS_lapSheet00();
  FSZZZZZZZ_patchSummary_(ss);

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    'Đã đối chiếu lại tổng chi trước VAT, VAT đầu vào, giá vốn, CPVH/bảo trì và Sheet tổng hợp.'
  );
}

function FSZZZZZZZ_rebuildSheet02TaxCost_(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  const idx = name => FSZZZZZZZ_headerIndex_(headers, name);

  const cV = idx('CP XD/TB trực tiếp trước VAT');
  const cW = idx('Chi phí bán hàng trước VAT');
  const cX = idx('Chi phí vận hành thuê trước VAT');
  const cY = idx('Chi phí GPMB phân bổ trước VAT');
  const cZ = idx('Chi phí HTKT phân bổ trước VAT');
  const cAA = idx('Tiền SDĐ/thuê đất phân bổ trước VAT');
  const cAB = idx('Chi phí dự phòng phân bổ trước VAT');
  const cAC = idx('Chi phí lãi vay phân bổ');
  const cAD = idx('Tổng giá vốn tính thuế');
  const cAE = idx('Lợi nhuận chịu thuế');
  const cAF = idx('Thuế TNDN tạm tính');
  const cQ = idx('Tổng doanh thu trước VAT');
  const cU = idx('Thuế TNDN %');

  // Các patch bảo trì có thể đổi tên cột X thành tổng CPVH & bảo trì.
  let cCombined = idx('Chi phí vận hành & bảo trì thuê trước VAT');
  if (cCombined < 0) cCombined = cX;

  const required = [cV, cW, cCombined, cY, cZ, cAA, cAB, cAC, cAD, cAE, cAF, cQ, cU];
  if (required.some(v => v < 0)) {
    throw new Error('Sheet 02 thiếu cột để đối chiếu giá vốn tính thuế.');
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  const values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

  const outAD = [], outAE = [], outAF = [];
  values.forEach(r => {
    const totalCost = [cV, cW, cCombined, cY, cZ, cAA, cAB, cAC]
      .reduce((s, c) => s + (Number(r[c]) || 0), 0);
    const taxable = Math.max(0, (Number(r[cQ]) || 0) - totalCost);
    const cit = taxable * (Number(r[cU]) || 0);
    outAD.push([totalCost]);
    outAE.push([taxable]);
    outAF.push([cit]);
  });

  sh.getRange(2, cAD + 1, outAD.length, 1).setValues(outAD);
  sh.getRange(2, cAE + 1, outAE.length, 1).setValues(outAE);
  sh.getRange(2, cAF + 1, outAF.length, 1).setValues(outAF);
}

function FSZZZZZZZ_rebuildSheet03Costs_(sh, vatRates, opVatRate) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  const col = name => FSZZZZZZZ_headerIndex_(headers, name);

  const cXD = col('Chi XD/TB/khác trước VAT');
  const cGPMB = col('Chi GPMB trước VAT');
  const cLand = col('Tiền SDĐ/thuê đất trước VAT');
  const cHTKT = col('Chi HTKT trước VAT');
  const cSelling = col('Chi phí bán hàng trước VAT');
  const cReserve = col('Chi phí dự phòng trước VAT');
  const cTotalBefore = col('Tổng chi trước VAT');
  const cVatIn = col('VAT đầu vào');
  const cTotalAfter = col('Tổng chi sau VAT');
  const cOp = col('Chi phí vận hành thuê trước VAT');
  const cMaint = col('Chi phí bảo trì trước VAT');

  const required = [cXD, cGPMB, cLand, cHTKT, cSelling, cReserve, cTotalBefore, cVatIn, cTotalAfter];
  if (required.some(v => v < 0)) throw new Error('Sheet 03 thiếu cột chi phí lõi.');
  if (cOp < 0 || cMaint < 0) throw new Error('Sheet 03 chưa có cột Chi phí vận hành/Chi phí bảo trì.');

  const rateXD = FSZZZZZZZ_pickRate_(vatRates, ['Chi phí XD/TB/khác']);
  const rateGPMB = FSZZZZZZZ_pickRate_(vatRates, ['Chi phí GPMB']);
  const rateLand = FSZZZZZZZ_pickRate_(vatRates, ['Tiền SDĐ', 'Tiền thuê đất']);
  const rateHTKT = FSZZZZZZZ_pickRate_(vatRates, ['Chi phí HTKT']);
  const rateSelling = FSZZZZZZZ_pickRate_(vatRates, ['Chi phí bán hàng']);
  const rateReserve = FSZZZZZZZ_pickRate_(vatRates, ['Chi phí dự phòng', 'Chi phí XD/TB/khác']);

  const values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  const outN = [], outO = [], outP = [];

  values.forEach(r => {
    const xd = Number(r[cXD]) || 0;
    const gpmb = Number(r[cGPMB]) || 0;
    const land = Number(r[cLand]) || 0;
    const htkt = Number(r[cHTKT]) || 0;
    const selling = Number(r[cSelling]) || 0;
    const reserve = Number(r[cReserve]) || 0;
    const op = Number(r[cOp]) || 0;
    const maint = Number(r[cMaint]) || 0;

    const totalBefore = xd + gpmb + land + htkt + selling + reserve + op + maint;
    const vatIn =
      xd * rateXD +
      gpmb * rateGPMB +
      land * rateLand +
      htkt * rateHTKT +
      selling * rateSelling +
      reserve * rateReserve +
      (op + maint) * opVatRate;

    outN.push([totalBefore]);
    outO.push([vatIn]);
    outP.push([totalBefore + vatIn]);
  });

  sh.getRange(2, cTotalBefore + 1, outN.length, 1).setValues(outN);
  sh.getRange(2, cVatIn + 1, outO.length, 1).setValues(outO);
  sh.getRange(2, cTotalAfter + 1, outP.length, 1).setValues(outP);
}

function FSZZZZZZZ_syncOperatingMaintenanceTo04_(ss) {
  const sh02 = ss.getSheetByName('02. Doanh thu');
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');
  if (!sh02 || !sh04) return;

  const h02 = sh02.getRange(1, 1, 1, sh02.getLastColumn()).getDisplayValues()[0];
  const cMonth = FSZZZZZZZ_headerIndex_(h02, 'Tháng số');
  const cOp = FSZZZZZZZ_headerIndex_(h02, 'Chi phí vận hành thuê trước VAT');
  const cMaint = FSZZZZZZZ_headerIndex_(h02, 'Chi phí bảo trì trước VAT');
  if (cMonth < 0 || cOp < 0 || cMaint < 0) return;

  const byMonth = {};
  const values = sh02.getRange(2, 1, Math.max(0, sh02.getLastRow() - 1), sh02.getLastColumn()).getValues();
  values.forEach(r => {
    const m = Number(r[cMonth]) || 0;
    if (!m) return;
    if (!byMonth[m]) byMonth[m] = { op: 0, maint: 0 };
    byMonth[m].op += Number(r[cOp]) || 0;
    byMonth[m].maint += Number(r[cMaint]) || 0;
  });

  const rowCount = Math.max(0, sh04.getLastRow() - 2);
  if (!rowCount) return;
  const months = sh04.getRange(3, 1, rowCount, 1).getValues();
  const op = [], maint = [], total = [];
  months.forEach(r => {
    const x = byMonth[Number(r[0]) || 0] || { op: 0, maint: 0 };
    op.push([x.op]);
    maint.push([x.maint]);
    total.push([x.op + x.maint]);
  });

  sh04.getRange(1, 40, 1, 3).merge().setValue('CHI PHÍ VẬN HÀNH & BẢO TRÌ');
  sh04.getRange(2, 40, 1, 3).setValues([[
    'Chi phí vận hành thuê trước VAT',
    'Chi phí bảo trì trước VAT',
    'Tổng chi phí vận hành & bảo trì trước VAT'
  ]]);
  sh04.getRange(3, 40, rowCount, 1).setValues(op);
  sh04.getRange(3, 41, rowCount, 1).setValues(maint);
  sh04.getRange(3, 42, rowCount, 1).setValues(total);
  sh04.getRange(3, 40, rowCount, 3).setNumberFormat('#,##0');
}

function FSZZZZZZZ_forceFCFE_(ss) {
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');
  if (!sh04) return;
  const n = Math.max(0, sh04.getLastRow() - 2);
  if (!n) return;
  sh04.getRange(3, 37, n, 1).setFormulaR1C1('=RC[-21]+RC[-15]-RC[-13]'); // AK=P+V-X
  sh04.getRange(3, 39, n, 1).setFormulaR1C1('=SUM(R3C37:RC37)');
}

function FSZZZZZZZ_patchSummary_(ss) {
  const sh = ss.getSheetByName('00. Tổng hợp');
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');
  if (!sh || !sh04) return;

  // Tổng CPVH và bảo trì sau VAT được cộng vào chỉ tiêu 2.
  const opMaintBefore = sh04.getRange(3, 42, Math.max(0, sh04.getLastRow() - 2), 1)
    .getValues().flat().reduce((s, v) => s + (Number(v) || 0), 0);

  const tech = ss.getSheetByName('01. Kỹ thuật');
  const vatRates = tech ? FSZZZZZZZ_readCommonCostVatRates_(tech) : {};
  const opVatRate = FSZZZZZZZ_pickRate_(vatRates, ['Chi phí vận hành', 'Chi phí bán hàng', 'Chi phí XD/TB/khác']);
  const opMaintAfter = opMaintBefore * (1 + opVatRate);

  // D30 hiện gồm Tổng vốn đầu tư + Chi phí bán hàng; cộng thêm CPVH/bảo trì.
  sh.getRange('D30').setFormula(`=SUM(D31:D32)+${opMaintAfter}/1000000000`);

  // Chỉ tiêu 1 = 2 + 3 + 12 + 13 theo yêu cầu đối chiếu.
  sh.getRange('D25').setFormula('=SUM(D30;D33;D42;D43)');
  sh.getRange('E25').setValue('Đối chiếu: Chỉ tiêu 1 = 2 + 3 + 12 + 13');
}

function FSZZZZZZZ_readCommonCostVatRates_(tech) {
  const data = tech.getDataRange().getValues();
  const out = {};
  let marker = -1;
  for (let r = 0; r < data.length; r++) {
    if (FSZZZZZZZ_norm_(data[r][0]) === 'chi phi chung') { marker = r; break; }
  }
  if (marker < 0) return out;

  for (let r = marker + 2; r < data.length; r++) {
    const name = String(data[r][0] || '').trim();
    if (!name || /^[A-Z_]+$/.test(name)) break;
    out[FSZZZZZZZ_norm_(name)] = Number(data[r][2]) || 0;
  }
  return out;
}

function FSZZZZZZZ_pickRate_(map, names) {
  for (const name of names) {
    const key = FSZZZZZZZ_norm_(name);
    if (Object.prototype.hasOwnProperty.call(map, key)) return Number(map[key]) || 0;
  }
  return 0;
}

function FSZZZZZZZ_headerIndex_(headers, name) {
  const target = FSZZZZZZZ_norm_(name);
  return headers.findIndex(h => FSZZZZZZZ_norm_(h) === target);
}

function FSZZZZZZZ_norm_(v) {
  return String(v || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}
