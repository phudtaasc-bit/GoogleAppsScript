/*************************************************
 * Bảo đảm Sheet 02 có đủ 2 cột vận hành và bảo trì
 * trước khi lập Sheet 00.
 * Nguồn lấy trực tiếp từ Sheet 03 theo Tháng số.
 *************************************************/

const FS_FIX_BASE_LAP_SHEET00_ = FS_lapSheet00;

FS_lapSheet00 = function() {
  FS_FIX_ensureSeparateOpMaintOnSheet02_();
  return FS_FIX_BASE_LAP_SHEET00_();
};

function FS_FIX_ensureSeparateOpMaintOnSheet02_() {
  const ss = SpreadsheetApp.getActive();
  const sh02 = ss.getSheetByName('02. Doanh thu');
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  if (!sh02) throw new Error('Không tìm thấy Sheet 02. Doanh thu.');
  if (!sh03) throw new Error('Không tìm thấy Sheet 03. Chi phí & Vốn.');

  const h02 = FS_FIX_detectHeader_(sh02, ['Tháng số']);
  const h03 = FS_FIX_detectHeader_(sh03, ['Tháng số']);

  let headers02 = sh02.getRange(h02, 1, 1, sh02.getLastColumn()).getDisplayValues()[0];
  const headers03 = sh03.getRange(h03, 1, 1, sh03.getLastColumn()).getDisplayValues()[0];

  const m02 = FS_FIX_findCol_(headers02, ['Tháng số']);
  const m03 = FS_FIX_findCol_(headers03, ['Tháng số']);

  let op02 = FS_FIX_findCol_(headers02, ['Chi phí vận hành thuê trước VAT', 'Chi phí vận hành trước VAT']);
  let mt02 = FS_FIX_findCol_(headers02, ['Chi phí bảo trì trước VAT']);

  const op03 = FS_FIX_findCol_(headers03, ['Chi phí vận hành thuê trước VAT', 'Chi phí vận hành trước VAT']);
  const mt03 = FS_FIX_findCol_(headers03, ['Chi phí bảo trì trước VAT']);

  if (m02 < 1 || m03 < 1) throw new Error('Không tìm thấy cột Tháng số tại Sheet 02 hoặc Sheet 03.');
  if (op03 < 1) throw new Error('Sheet 03 thiếu cột Chi phí vận hành thuê trước VAT.');
  if (mt03 < 1) throw new Error('Sheet 03 thiếu cột Chi phí bảo trì trước VAT.');

  if (op02 < 1) {
    op02 = sh02.getLastColumn() + 1;
    sh02.getRange(h02, op02).setValue('Chi phí vận hành thuê trước VAT');
  }
  if (mt02 < 1) {
    mt02 = sh02.getLastColumn() + 1;
    sh02.getRange(h02, mt02).setValue('Chi phí bảo trì trước VAT');
  }

  const start02 = h02 + 1;
  const start03 = h03 + 1;
  const n02 = Math.max(0, sh02.getLastRow() - h02);
  const n03 = Math.max(0, sh03.getLastRow() - h03);
  if (!n02) return;

  const data03 = n03 ? sh03.getRange(start03, 1, n03, sh03.getLastColumn()).getValues() : [];
  const byMonth = new Map();
  data03.forEach(r => {
    const month = Number(r[m03 - 1]) || 0;
    if (!month) return;
    byMonth.set(month, {
      op: Number(r[op03 - 1]) || 0,
      mt: Number(r[mt03 - 1]) || 0
    });
  });

  const months02 = sh02.getRange(start02, m02, n02, 1).getValues();
  const outOp = [];
  const outMt = [];
  months02.forEach(r => {
    const month = Number(r[0]) || 0;
    const src = byMonth.get(month) || { op: 0, mt: 0 };
    outOp.push([month ? src.op : '']);
    outMt.push([month ? src.mt : '']);
  });

  sh02.getRange(start02, op02, n02, 1).setValues(outOp);
  sh02.getRange(start02, mt02, n02, 1).setValues(outMt);
  SpreadsheetApp.flush();
}

function FS_FIX_detectHeader_(sh, requiredNames) {
  const maxRows = Math.min(6, sh.getLastRow());
  let bestRow = -1;
  let bestHits = -1;
  for (let r = 1; r <= maxRows; r++) {
    const headers = sh.getRange(r, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
    const hits = requiredNames.reduce((s, n) => s + (FS_FIX_findCol_(headers, [n]) > 0 ? 1 : 0), 0);
    if (hits > bestHits) {
      bestHits = hits;
      bestRow = r;
    }
  }
  if (bestRow < 1 || bestHits < 1) throw new Error('Không nhận diện được dòng tiêu đề tại ' + sh.getName() + '.');
  return bestRow;
}

function FS_FIX_findCol_(headers, names) {
  const normalized = headers.map(FS_FIX_key_);
  for (const name of names) {
    const idx = normalized.indexOf(FS_FIX_key_(name));
    if (idx >= 0) return idx + 1;
  }
  return -1;
}

function FS_FIX_key_(v) {
  return String(v || '').toLowerCase().normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}
