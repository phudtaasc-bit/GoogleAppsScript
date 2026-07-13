/*************************************************
 * ZZZZZZ_CostReserveOpsFinal.js
 *
 * Mục tiêu:
 * 1) Chi phí dự phòng theo tháng = Tỷ lệ dự phòng × (Chi XD/TB/khác + Chi HTKT)
 *    phát sinh đúng tại tháng có chi phí XD/HTKT.
 * 2) Tổng chi phí và dòng tiền được cập nhật sau dự phòng, nên nhu cầu vốn và lãi vay
 *    tự động tính trên dòng tiền đã bao gồm dự phòng.
 * 3) Đồng bộ chi phí vận hành và bảo trì từ Sheet 02 sang Sheet 03 và Sheet 04.
 *
 * Quy ước cột:
 * Sheet 02: A tháng số; AH CP vận hành; AI CP bảo trì.
 * Sheet 03: H XD/TB; K HTKT; M dự phòng; N tổng chi trước VAT;
 *           O VAT đầu vào; P tổng chi sau VAT; AJ vận hành; AK bảo trì; AL tổng VH+BT.
 * Sheet 04: AN vận hành; AO bảo trì; AP tổng VH+BT.
 *************************************************/

function FSZZZZZZ_capNhatDuPhongVaVanHanhBaoTri() {
  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName('01. Kỹ thuật');
  const sh02 = ss.getSheetByName('02. Doanh thu');
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');

  if (!tech || !sh02 || !sh03 || !sh04) {
    throw new Error('Thiếu Sheet 01. Kỹ thuật, 02, 03 hoặc 04.');
  }

  const reserveCfg = FSZZZZZZ_getReserveConfig_(tech);
  const reserveRate = reserveCfg.rate;
  const reserveVatRate = reserveCfg.vatRate;

  const opMaintByMonth = FSZZZZZZ_getOpMaintByMonth_(sh02);

  const lastRow03 = sh03.getLastRow();
  if (lastRow03 < 2) return;

  const rowCount03 = lastRow03 - 1;
  const width03 = Math.max(38, sh03.getLastColumn());
  const data03 = sh03.getRange(2, 1, rowCount03, width03).getValues();

  const outM = [];
  const outN = [];
  const outO = [];
  const outP = [];
  const outAJ = [];
  const outAK = [];
  const outAL = [];

  data03.forEach(row => {
    const month = Number(row[0]) || 0;       // A
    const xd = Number(row[7]) || 0;          // H
    const htkt = Number(row[10]) || 0;       // K

    const oldReserve = Number(row[12]) || 0; // M
    const oldTotalPreVat = Number(row[13]) || 0; // N
    const oldVatIn = Number(row[14]) || 0;   // O
    const oldTotalAfterVat = Number(row[15]) || 0; // P

    const oldOp = Number(row[35]) || 0;      // AJ
    const oldMaint = Number(row[36]) || 0;   // AK
    const oldOpMaint = Number(row[37]) || (oldOp + oldMaint); // AL

    const newReserve = Math.max(0, (xd + htkt) * reserveRate);
    const x = opMaintByMonth[month] || { op: 0, maint: 0 };
    const newOp = Math.max(0, Number(x.op) || 0);
    const newMaint = Math.max(0, Number(x.maint) || 0);
    const newOpMaint = newOp + newMaint;

    // Điều chỉnh theo chênh lệch để hàm chạy lặp vẫn không cộng trùng.
    const reserveDelta = newReserve - oldReserve;
    const opMaintDelta = newOpMaint - oldOpMaint;
    const reserveVatDelta = reserveDelta * reserveVatRate;

    const newTotalPreVat = oldTotalPreVat + reserveDelta + opMaintDelta;
    const newVatIn = oldVatIn + reserveVatDelta;
    const newTotalAfterVat = oldTotalAfterVat + reserveDelta + opMaintDelta + reserveVatDelta;

    outM.push([newReserve]);
    outN.push([newTotalPreVat]);
    outO.push([newVatIn]);
    outP.push([newTotalAfterVat]);
    outAJ.push([newOp]);
    outAK.push([newMaint]);
    outAL.push([newOpMaint]);
  });

  sh03.getRange(2, 13, rowCount03, 1).setValues(outM);  // M
  sh03.getRange(2, 14, rowCount03, 1).setValues(outN);  // N
  sh03.getRange(2, 15, rowCount03, 1).setValues(outO);  // O
  sh03.getRange(2, 16, rowCount03, 1).setValues(outP);  // P
  sh03.getRange(2, 36, rowCount03, 1).setValues(outAJ); // AJ
  sh03.getRange(2, 37, rowCount03, 1).setValues(outAK); // AK
  sh03.getRange(2, 38, rowCount03, 1).setValues(outAL); // AL

  sh03.getRange(1, 36).setValue('Chi phí vận hành thuê trước VAT');
  sh03.getRange(1, 37).setValue('Chi phí bảo trì trước VAT');
  sh03.getRange(1, 38).setValue('Tổng chi phí vận hành & bảo trì trước VAT');

  SpreadsheetApp.flush();

  // Lập lại Sheet 04 để dòng tiền trước tài trợ, nhu cầu vốn, giải ngân và lãi vay
  // sử dụng tổng chi phí sau khi đã cập nhật dự phòng và vận hành/bảo trì.
  FS_lapSheet04();

  const lastRow04 = sh04.getLastRow();
  if (lastRow04 >= 3) {
    const rowCount04 = lastRow04 - 2;
    const months04 = sh04.getRange(3, 1, rowCount04, 1).getValues();
    const outAN = [];
    const outAO = [];
    const outAP = [];

    months04.forEach(r => {
      const month = Number(r[0]) || 0;
      const x = opMaintByMonth[month] || { op: 0, maint: 0 };
      const op = Math.max(0, Number(x.op) || 0);
      const maint = Math.max(0, Number(x.maint) || 0);
      outAN.push([op]);
      outAO.push([maint]);
      outAP.push([op + maint]);
    });

    sh04.getRange(3, 40, rowCount04, 1).setValues(outAN); // AN
    sh04.getRange(3, 41, rowCount04, 1).setValues(outAO); // AO
    sh04.getRange(3, 42, rowCount04, 1).setValues(outAP); // AP

    sh04.getRange(1, 40, 1, 3).setValues([['CHI PHÍ VẬN HÀNH & BẢO TRÌ', '', '']]);
    sh04.getRange(2, 40, 1, 3).setValues([[
      'Chi phí vận hành thuê trước VAT',
      'Chi phí bảo trì trước VAT',
      'Tổng chi phí vận hành & bảo trì trước VAT'
    ]]);
  }

  SpreadsheetApp.flush();
  ss.toast('Đã cập nhật dự phòng, vận hành và bảo trì; Sheet 04 đã được lập lại.', 'FS', 8);
}

function FSZZZZZZ_getReserveConfig_(tech) {
  const values = tech.getDataRange().getValues();
  const display = tech.getDataRange().getDisplayValues();

  let inCostTable = false;
  let headerRow = -1;
  let rateCol = -1;
  let vatCol = -1;

  for (let r = 0; r < display.length; r++) {
    const rowText = display[r].map(v => String(v || '').trim());
    if (rowText.some(v => FSZZZZZZ_norm_(v) === 'chi phi chung')) {
      inCostTable = true;
      continue;
    }

    if (inCostTable && headerRow < 0) {
      const keys = rowText.map(FSZZZZZZ_norm_);
      if (keys.includes('khoan muc')) {
        headerRow = r;
        rateCol = keys.indexOf('ty le');
        vatCol = keys.indexOf('vat dau vao');
        continue;
      }
    }

    if (headerRow >= 0) {
      const item = FSZZZZZZ_norm_(rowText[0]);
      if (item === 'chi phi du phong' || item === 'du phong') {
        const rawRate = rateCol >= 0 ? values[r][rateCol] : 0;
        const rawVat = vatCol >= 0 ? values[r][vatCol] : 0;
        return {
          rate: FSZZZZZZ_rate_(rawRate),
          vatRate: FSZZZZZZ_rate_(rawVat)
        };
      }

      if (rowText.some(v => /^san_pham$/i.test(String(v)))) break;
    }
  }

  return { rate: 0, vatRate: 0 };
}

function FSZZZZZZ_getOpMaintByMonth_(sh02) {
  const out = {};
  const lastRow = sh02.getLastRow();
  if (lastRow < 2) return out;

  const lastCol = Math.max(35, sh02.getLastColumn());
  const rows = sh02.getRange(2, 1, lastRow - 1, lastCol).getValues();

  rows.forEach(r => {
    const month = Number(r[0]) || 0;
    if (!month) return;

    const op = Number(r[33]) || 0;    // AH
    const maint = Number(r[34]) || 0; // AI

    if (!out[month]) out[month] = { op: 0, maint: 0 };
    out[month].op += op;
    out[month].maint += maint;
  });

  return out;
}

function FSZZZZZZ_rate_(v) {
  if (typeof v === 'number') {
    if (!isFinite(v)) return 0;
    return Math.abs(v) > 1 ? v / 100 : v;
  }

  const s = String(v || '').trim().replace(/\s/g, '').replace(',', '.');
  if (!s) return 0;
  const n = Number(s.replace('%', ''));
  if (!isFinite(n)) return 0;
  return s.includes('%') || Math.abs(n) > 1 ? n / 100 : n;
}

function FSZZZZZZ_norm_(v) {
  return String(v || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/[_\s]+/g, ' ')
    .trim();
}
