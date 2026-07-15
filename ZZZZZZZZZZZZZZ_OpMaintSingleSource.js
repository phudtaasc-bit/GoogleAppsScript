/*************************************************
 * ZZZZZZZZZZZZZZ_OpMaintSingleSource.js
 * Chuẩn hóa chi phí vận hành/bảo trì tại Sheet 02:
 * - AH: Chi phí vận hành thuê trước VAT
 * - AI: Chi phí bảo trì trước VAT
 * - X : Chi phí vận hành & bảo trì thuê trước VAT = AH + AI
 *
 * Không tính lại AH/AI. Cột X chỉ là cột tổng dẫn xuất.
 *************************************************/

const FS_OPM_BASE_LAP_SHEET02_ = FS_lapSheet02;

FS_lapSheet02 = function() {
  FS_OPM_BASE_LAP_SHEET02_();
  FS_OPM_syncCombinedFromSeparate_();
};

function FS_OPM_syncCombinedFromSeparate_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('02. Doanh thu');
  if (!sh) throw new Error('Không tìm thấy Sheet 02. Doanh thu.');

  const headerRow = FS_OPM_detectHeaderRow_(sh);
  const headers = sh.getRange(headerRow, 1, 1, sh.getLastColumn()).getDisplayValues()[0];

  const monthCol = FS_OPM_findCol_(headers, ['Tháng số']);
  const combinedCol = FS_OPM_findCol_(headers, [
    'Chi phí vận hành & bảo trì thuê trước VAT',
    'Chi phí vận hành và bảo trì thuê trước VAT'
  ]);
  const opCol = FS_OPM_findCol_(headers, [
    'Chi phí vận hành thuê trước VAT',
    'Chi phí vận hành trước VAT'
  ]);
  const maintCol = FS_OPM_findCol_(headers, ['Chi phí bảo trì trước VAT']);

  if (monthCol < 1) throw new Error('Sheet 02 thiếu cột Tháng số.');
  if (opCol < 1) throw new Error('Sheet 02 thiếu cột Chi phí vận hành thuê trước VAT.');
  if (maintCol < 1) throw new Error('Sheet 02 thiếu cột Chi phí bảo trì trước VAT.');
  if (combinedCol < 1) throw new Error('Sheet 02 thiếu cột Chi phí vận hành & bảo trì thuê trước VAT.');

  const startRow = headerRow + 1;
  const n = sh.getLastRow() - headerRow;
  if (n <= 0) return;

  const months = sh.getRange(startRow, monthCol, n, 1).getValues();
  const op = sh.getRange(startRow, opCol, n, 1).getValues();
  const maint = sh.getRange(startRow, maintCol, n, 1).getValues();

  const out = [];
  for (let i = 0; i < n; i++) {
    const month = Number(months[i][0]) || 0;
    if (!month) {
      out.push(['']);
      continue;
    }
    out.push([(Number(op[i][0]) || 0) + (Number(maint[i][0]) || 0)]);
  }

  sh.getRange(startRow, combinedCol, n, 1)
    .setValues(out)
    .setNumberFormat('#,##0');
  SpreadsheetApp.flush();
}

function FS_OPM_detectHeaderRow_(sh) {
  const maxRows = Math.min(8, sh.getLastRow());
  let bestRow = 1;
  let bestHits = -1;
  for (let r = 1; r <= maxRows; r++) {
    const headers = sh.getRange(r, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
    let hits = 0;
    if (FS_OPM_findCol_(headers, ['Tháng số']) > 0) hits++;
    if (FS_OPM_findCol_(headers, ['Chi phí vận hành thuê trước VAT', 'Chi phí vận hành trước VAT']) > 0) hits++;
    if (FS_OPM_findCol_(headers, ['Chi phí bảo trì trước VAT']) > 0) hits++;
    if (FS_OPM_findCol_(headers, ['Chi phí vận hành & bảo trì thuê trước VAT', 'Chi phí vận hành và bảo trì thuê trước VAT']) > 0) hits++;
    if (hits > bestHits) {
      bestHits = hits;
      bestRow = r;
    }
  }
  if (bestHits < 3) throw new Error('Không nhận diện được dòng tiêu đề Sheet 02 cho chi phí vận hành/bảo trì.');
  return bestRow;
}

function FS_OPM_findCol_(headers, aliases) {
  const normalized = headers.map(FS_OPM_key_);
  for (const alias of aliases) {
    const key = FS_OPM_key_(alias);
    const idx = normalized.indexOf(key);
    if (idx >= 0) return idx + 1;
  }
  return -1;
}

function FS_OPM_key_(v) {
  return String(v || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/&/g, ' va ')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}
