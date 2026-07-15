/*************************************************
 * Fix VAT đầu vào Sheet 03:
 * - VAT đầu vào phải bao gồm VAT vận hành và VAT bảo trì.
 * - Tổng chi trước VAT phải bao gồm vận hành và bảo trì.
 * - Tổng chi sau VAT = Tổng chi trước VAT + VAT đầu vào.
 *
 * Patch bọc hàm rebuild hiện hữu, không thay đổi các logic VAT khác.
 *************************************************/

const FS03_VAT_BASE_REBUILD_COSTS_ = FSZZZZZZZ_rebuildSheet03Costs_;

FSZZZZZZZ_rebuildSheet03Costs_ = function(sh03, vatRates, opVatRate) {
  const result = FS03_VAT_BASE_REBUILD_COSTS_(sh03, vatRates, opVatRate);
  FS03_VAT_includeOperatingMaintenance_(sh03, opVatRate);
  return result;
};

function FS03_VAT_includeOperatingMaintenance_(sh03, opVatRate) {
  if (!sh03) throw new Error('Thiếu Sheet 03 để cập nhật VAT vận hành/bảo trì.');

  const headerRow = FS03_VAT_detectHeaderRow_(sh03);
  const lastCol = sh03.getLastColumn();
  const headers = sh03.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];

  const colMonth = FS03_VAT_findCol_(headers, ['Tháng số']);
  const colOp = FS03_VAT_findCol_(headers, [
    'Chi phí vận hành thuê trước VAT',
    'Chi phí vận hành trước VAT'
  ]);
  const colMaint = FS03_VAT_findCol_(headers, ['Chi phí bảo trì trước VAT']);
  const colTotalBefore = FS03_VAT_findCol_(headers, ['Tổng chi trước VAT']);
  const colInputVat = FS03_VAT_findCol_(headers, ['VAT đầu vào']);
  const colTotalAfter = FS03_VAT_findCol_(headers, ['Tổng chi sau VAT']);

  const required = [
    ['Tháng số', colMonth],
    ['Chi phí vận hành', colOp],
    ['Chi phí bảo trì', colMaint],
    ['Tổng chi trước VAT', colTotalBefore],
    ['VAT đầu vào', colInputVat],
    ['Tổng chi sau VAT', colTotalAfter]
  ];
  const missing = required.filter(x => x[1] < 1).map(x => x[0]);
  if (missing.length) {
    throw new Error('Sheet 03 thiếu cột để tính VAT vận hành/bảo trì: ' + missing.join(', ') + '.');
  }

  const rate = Number(opVatRate);
  if (!isFinite(rate) || rate < 0) {
    throw new Error('Thuế suất VAT vận hành/bảo trì không hợp lệ: ' + opVatRate);
  }

  const startRow = headerRow + 1;
  const n = Math.max(0, sh03.getLastRow() - headerRow);
  if (!n) return;

  const months = sh03.getRange(startRow, colMonth, n, 1).getValues();
  const opVals = sh03.getRange(startRow, colOp, n, 1).getValues();
  const maintVals = sh03.getRange(startRow, colMaint, n, 1).getValues();
  const beforeVals = sh03.getRange(startRow, colTotalBefore, n, 1).getValues();
  const vatVals = sh03.getRange(startRow, colInputVat, n, 1).getValues();

  const outBefore = [];
  const outVat = [];
  const outAfter = [];

  for (let i = 0; i < n; i++) {
    const month = Number(months[i][0]) || 0;
    if (!month) {
      outBefore.push(['']);
      outVat.push(['']);
      outAfter.push(['']);
      continue;
    }

    const op = Number(opVals[i][0]) || 0;
    const maint = Number(maintVals[i][0]) || 0;
    const oldBefore = Number(beforeVals[i][0]) || 0;
    const oldVat = Number(vatVals[i][0]) || 0;

    // Hàm cũ chưa cộng vận hành/bảo trì vào tổng chi trước VAT.
    const totalBefore = oldBefore + op + maint;
    const inputVat = oldVat + (op + maint) * rate;
    const totalAfter = totalBefore + inputVat;

    outBefore.push([totalBefore]);
    outVat.push([inputVat]);
    outAfter.push([totalAfter]);
  }

  sh03.getRange(startRow, colTotalBefore, n, 1).setValues(outBefore).setNumberFormat('#,##0');
  sh03.getRange(startRow, colInputVat, n, 1).setValues(outVat).setNumberFormat('#,##0');
  sh03.getRange(startRow, colTotalAfter, n, 1).setValues(outAfter).setNumberFormat('#,##0');
  SpreadsheetApp.flush();
}

function FS03_VAT_detectHeaderRow_(sh) {
  const maxRows = Math.min(8, sh.getLastRow());
  let bestRow = 1;
  let bestHits = -1;

  for (let r = 1; r <= maxRows; r++) {
    const headers = sh.getRange(r, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
    let hits = 0;
    if (FS03_VAT_findCol_(headers, ['Tháng số']) > 0) hits++;
    if (FS03_VAT_findCol_(headers, ['VAT đầu vào']) > 0) hits++;
    if (FS03_VAT_findCol_(headers, ['Tổng chi trước VAT']) > 0) hits++;
    if (FS03_VAT_findCol_(headers, ['Chi phí vận hành thuê trước VAT', 'Chi phí vận hành trước VAT']) > 0) hits++;
    if (FS03_VAT_findCol_(headers, ['Chi phí bảo trì trước VAT']) > 0) hits++;

    if (hits > bestHits) {
      bestHits = hits;
      bestRow = r;
    }
  }

  if (bestHits < 4) throw new Error('Không nhận diện được dòng tiêu đề Sheet 03 để cập nhật VAT.');
  return bestRow;
}

function FS03_VAT_findCol_(headers, aliases) {
  const normalized = headers.map(FS03_VAT_key_);
  for (const alias of aliases) {
    const idx = normalized.indexOf(FS03_VAT_key_(alias));
    if (idx >= 0) return idx + 1;
  }
  return -1;
}

function FS03_VAT_key_(v) {
  return String(v || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}
