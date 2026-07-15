/*************************************************
 * CoreCostRebuildV3
 * Chuẩn hóa trực tiếp Sheet 03 sau khi hàm lập Sheet 03 hiện hữu chạy xong.
 *
 * Nguyên tắc:
 * 1) Chi phí dự phòng tháng = Tỷ lệ dự phòng × (Chi XD/TB tháng + Chi HTKT tháng).
 * 2) Chi phí vận hành/bảo trì tại Sheet 03 lấy tổng theo Tháng số từ Sheet 02.
 * 3) VAT đầu vào cộng VAT dự phòng, vận hành và bảo trì đúng một lần.
 * 4) Tổng chi trước/sau VAT và dòng tiền trước tài trợ được tính lại tuyệt đối.
 *
 * File này là lớp chuẩn hóa cuối cùng, không cộng dồn trên số đã vá qua nhiều vòng.
 *************************************************/

const FS_COREV3_BASE_LAP_SHEET03_ = FS_lapSheet03;

FS_lapSheet03 = function() {
  const result = FS_COREV3_BASE_LAP_SHEET03_();
  FS_COREV3_normalizeSheet03_();
  return result;
};

function FS_COREV3_normalizeSheet03_() {
  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName('01. Kỹ thuật');
  const sh02 = ss.getSheetByName('02. Doanh thu');
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');

  if (!tech || !sh02 || !sh03) {
    throw new Error('Thiếu Sheet 01. Kỹ thuật, 02. Doanh thu hoặc 03. Chi phí & Vốn.');
  }

  const h02 = FS_COREV3_detectHeader_(sh02, ['Tháng số']);
  const h03 = FS_COREV3_detectHeader_(sh03, ['Tháng số', 'VAT đầu vào', 'Tổng chi trước VAT']);

  const headers02 = sh02.getRange(h02, 1, 1, sh02.getLastColumn()).getDisplayValues()[0];
  let headers03 = sh03.getRange(h03, 1, 1, sh03.getLastColumn()).getDisplayValues()[0];

  const c02Month = FS_COREV3_col_(headers02, ['Tháng số']);
  const c02Op = FS_COREV3_col_(headers02, ['Chi phí vận hành thuê trước VAT', 'Chi phí vận hành trước VAT']);
  const c02Maint = FS_COREV3_col_(headers02, ['Chi phí bảo trì trước VAT']);

  if (c02Month < 1 || c02Op < 1 || c02Maint < 1) {
    throw new Error('Sheet 02 thiếu cột Tháng số, Chi phí vận hành hoặc Chi phí bảo trì.');
  }

  let c03Op = FS_COREV3_col_(headers03, ['Chi phí vận hành thuê trước VAT', 'Chi phí vận hành trước VAT']);
  let c03Maint = FS_COREV3_col_(headers03, ['Chi phí bảo trì trước VAT']);
  let c03Combined = FS_COREV3_col_(headers03, ['Tổng chi phí vận hành & bảo trì trước VAT', 'Tổng chi phí vận hành và bảo trì trước VAT']);

  const requiredExtra = [
    ['Chi phí vận hành thuê trước VAT', c03Op],
    ['Chi phí bảo trì trước VAT', c03Maint],
    ['Tổng chi phí vận hành & bảo trì trước VAT', c03Combined]
  ];
  requiredExtra.forEach(item => {
    if (item[1] < 1) {
      const newCol = sh03.getLastColumn() + 1;
      sh03.getRange(h03, newCol).setValue(item[0]);
      if (item[0].indexOf('vận hành thuê') >= 0) c03Op = newCol;
      else if (item[0] === 'Chi phí bảo trì trước VAT') c03Maint = newCol;
      else c03Combined = newCol;
    }
  });

  headers03 = sh03.getRange(h03, 1, 1, sh03.getLastColumn()).getDisplayValues()[0];

  const cMonth = FS_COREV3_requireCol_(headers03, ['Tháng số']);
  const cCashKH = FS_COREV3_requireCol_(headers03, ['Dòng tiền huy động từ KH']);
  const cVatOut = FS_COREV3_requireCol_(headers03, ['VAT đầu ra']);
  const cCit = FS_COREV3_requireCol_(headers03, ['Thuế TNDN']);
  const cXD = FS_COREV3_requireCol_(headers03, ['Chi XD/TB/khác trước VAT', 'Chi phí XD/TB/khác trước VAT']);
  const cGPMB = FS_COREV3_requireCol_(headers03, ['Chi GPMB trước VAT', 'Chi phí GPMB trước VAT']);
  const cLand = FS_COREV3_requireCol_(headers03, ['Tiền SDĐ/thuê đất trước VAT']);
  const cHTKT = FS_COREV3_requireCol_(headers03, ['Chi HTKT trước VAT', 'Chi phí HTKT trước VAT']);
  const cSelling = FS_COREV3_requireCol_(headers03, ['Chi phí bán hàng trước VAT', 'Chi bán hàng trước VAT']);
  const cReserve = FS_COREV3_requireCol_(headers03, ['Chi phí dự phòng trước VAT']);
  const cTotalBefore = FS_COREV3_requireCol_(headers03, ['Tổng chi trước VAT']);
  const cInputVat = FS_COREV3_requireCol_(headers03, ['VAT đầu vào']);
  const cTotalAfter = FS_COREV3_requireCol_(headers03, ['Tổng chi sau VAT']);
  const cVatCarryIn = FS_COREV3_requireCol_(headers03, ['VAT còn được khấu trừ đầu kỳ']);
  const cVatPayable = FS_COREV3_requireCol_(headers03, ['VAT phải nộp']);
  const cVatCarryOut = FS_COREV3_requireCol_(headers03, ['VAT còn được khấu trừ cuối kỳ']);
  const cCashBeforeFunding = FS_COREV3_requireCol_(headers03, ['Dòng tiền trước tài trợ']);

  const commonCosts = FS03V21_docChiPhiChung_(tech);
  const reserveRate = FS03V21_layTyLe_(commonCosts, ['Chi phí dự phòng']);
  const reserveVatRate = FS03V21_layVAT_(commonCosts, ['Chi phí dự phòng']);
  const opVatRate = FS_COREV3_firstRate_(commonCosts, [
    'Chi phí vận hành', 'Chi phí vận hành thuê', 'Chi phí vận hành & bảo trì'
  ]);
  const maintVatRateRaw = FS_COREV3_firstRate_(commonCosts, ['Chi phí bảo trì']);
  const maintVatRate = maintVatRateRaw >= 0 ? maintVatRateRaw : opVatRate;

  const start02 = h02 + 1;
  const n02 = Math.max(0, sh02.getLastRow() - h02);
  const monthMap = new Map();
  if (n02 > 0) {
    const data02 = sh02.getRange(start02, 1, n02, sh02.getLastColumn()).getValues();
    data02.forEach(r => {
      const month = Number(r[c02Month - 1]) || 0;
      if (!month) return;
      const item = monthMap.get(month) || { op: 0, maint: 0 };
      item.op += Number(r[c02Op - 1]) || 0;
      item.maint += Number(r[c02Maint - 1]) || 0;
      monthMap.set(month, item);
    });
  }

  const start03 = h03 + 1;
  const n03 = Math.max(0, sh03.getLastRow() - h03);
  if (!n03) return;

  const data03 = sh03.getRange(start03, 1, n03, sh03.getLastColumn()).getValues();
  const out = data03.map(r => r.slice());

  let carryIn = 0;
  for (let i = 0; i < out.length; i++) {
    const r = out[i];
    const month = Number(r[cMonth - 1]) || 0;
    if (!month) continue;

    const xd = Number(r[cXD - 1]) || 0;
    const gpmb = Number(r[cGPMB - 1]) || 0;
    const land = Number(r[cLand - 1]) || 0;
    const htkt = Number(r[cHTKT - 1]) || 0;
    const selling = Number(r[cSelling - 1]) || 0;
    const oldReserve = Number(r[cReserve - 1]) || 0;
    const oldInputVat = Number(r[cInputVat - 1]) || 0;
    const item = monthMap.get(month) || { op: 0, maint: 0 };

    const reserve = reserveRate * (xd + htkt);
    const op = item.op;
    const maint = item.maint;
    const combined = op + maint;

    // Hàm gốc đã tính VAT cho các cấu phần cũ, gồm dự phòng cũ.
    // Loại VAT dự phòng cũ rồi cộng lại VAT dự phòng mới và VAT vận hành/bảo trì.
    const baseVatExcludingReserve = Math.max(0, oldInputVat - oldReserve * reserveVatRate);
    const inputVat = baseVatExcludingReserve
      + reserve * reserveVatRate
      + op * opVatRate
      + maint * maintVatRate;

    const totalBefore = xd + gpmb + land + htkt + selling + reserve + op + maint;
    const totalAfter = totalBefore + inputVat;
    const vatOut = Number(r[cVatOut - 1]) || 0;
    const cit = Number(r[cCit - 1]) || 0;
    const cashKH = Number(r[cCashKH - 1]) || 0;
    const vatPayable = Math.max(0, vatOut - carryIn - inputVat);
    const carryOut = Math.max(0, carryIn + inputVat - vatOut);
    const cashBeforeFunding = cashKH - totalAfter - vatPayable - cit;

    r[cReserve - 1] = reserve;
    r[c03Op - 1] = op;
    r[c03Maint - 1] = maint;
    r[c03Combined - 1] = combined;
    r[cTotalBefore - 1] = totalBefore;
    r[cInputVat - 1] = inputVat;
    r[cTotalAfter - 1] = totalAfter;
    r[cVatCarryIn - 1] = carryIn;
    r[cVatPayable - 1] = vatPayable;
    r[cVatCarryOut - 1] = carryOut;
    r[cCashBeforeFunding - 1] = cashBeforeFunding;

    carryIn = carryOut;
  }

  sh03.getRange(start03, 1, out.length, out[0].length).setValues(out);
  sh03.getRange(start03, cReserve, n03, 1).setNumberFormat('#,##0');
  sh03.getRange(start03, cInputVat, n03, 1).setNumberFormat('#,##0');
  sh03.getRange(start03, c03Op, n03, 3).setNumberFormat('#,##0');
  SpreadsheetApp.flush();
}

function FS_COREV3_firstRate_(commonCosts, aliases) {
  for (const alias of aliases) {
    const value = FS03V21_layVAT_(commonCosts, [alias]);
    if (isFinite(value) && value >= 0) return Number(value) || 0;
  }
  return 0;
}

function FS_COREV3_detectHeader_(sh, required) {
  const maxRows = Math.min(8, sh.getLastRow());
  let bestRow = 1;
  let bestHits = -1;
  for (let r = 1; r <= maxRows; r++) {
    const headers = sh.getRange(r, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
    const hits = required.reduce((s, name) => s + (FS_COREV3_col_(headers, [name]) > 0 ? 1 : 0), 0);
    if (hits > bestHits) {
      bestHits = hits;
      bestRow = r;
    }
  }
  if (bestHits < 1) throw new Error('Không nhận diện được dòng tiêu đề tại ' + sh.getName() + '.');
  return bestRow;
}

function FS_COREV3_requireCol_(headers, aliases) {
  const col = FS_COREV3_col_(headers, aliases);
  if (col < 1) throw new Error('Không tìm thấy cột: ' + aliases.join(' / '));
  return col;
}

function FS_COREV3_col_(headers, aliases) {
  const normalized = headers.map(FS_COREV3_key_);
  for (const alias of aliases) {
    const idx = normalized.indexOf(FS_COREV3_key_(alias));
    if (idx >= 0) return idx + 1;
  }
  return -1;
}

function FS_COREV3_key_(v) {
  return String(v || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/&/g, ' va ')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}
