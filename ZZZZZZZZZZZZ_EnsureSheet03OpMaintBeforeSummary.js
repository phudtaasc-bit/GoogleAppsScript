/*************************************************
 * ZZZZZZZZZZZZ_EnsureSheet03OpMaintBeforeSummary.js
 * Khôi phục chắc chắn 2 cột vận hành/bảo trì tại Sheet 03.
 * Áp dụng ngay sau khi lập Sheet 03 và trước khi lập Sheet 00.
 *************************************************/

const FS03_OPMAINT_BASE_LAP03_ = FS_lapSheet03;
FS_lapSheet03 = function() {
  const result = FS03_OPMAINT_BASE_LAP03_.apply(this, arguments);
  FS03_OPMAINT_ensureFromSheet02_();
  return result;
};

const FS03_OPMAINT_BASE_LAP00_ = FS_lapSheet00;
FS_lapSheet00 = function() {
  FS03_OPMAINT_ensureFromSheet02_();
  return FS03_OPMAINT_BASE_LAP00_.apply(this, arguments);
};

function FS03_OPMAINT_ensureFromSheet02_() {
  const ss = SpreadsheetApp.getActive();
  const sh02 = ss.getSheetByName('02. Doanh thu');
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  if (!sh02 || !sh03) return;

  const h02 = sh02.getRange(1, 1, 1, sh02.getLastColumn()).getDisplayValues()[0];
  let h03 = sh03.getRange(1, 1, 1, sh03.getLastColumn()).getDisplayValues()[0];

  const month02 = FS03_OPMAINT_find_(h02, ['Tháng số']);
  const month03 = FS03_OPMAINT_find_(h03, ['Tháng số']);
  if (month02 < 0 || month03 < 0) {
    throw new Error('Không nhận diện được cột Tháng số tại Sheet 02/03.');
  }

  const op02 = FS03_OPMAINT_find_(h02, [
    'Chi phí vận hành thuê trước VAT',
    'Chi phí vận hành trước VAT'
  ]);
  const maint02 = FS03_OPMAINT_find_(h02, ['Chi phí bảo trì trước VAT']);
  const combined02 = FS03_OPMAINT_find_(h02, [
    'Chi phí vận hành & bảo trì thuê trước VAT',
    'Chi phí vận hành và bảo trì thuê trước VAT',
    'Tổng chi phí vận hành & bảo trì trước VAT',
    'Tổng chi phí vận hành và bảo trì trước VAT'
  ]);

  if (op02 < 0 && maint02 < 0 && combined02 < 0) {
    throw new Error('Sheet 02 không có cột vận hành/bảo trì để đồng bộ sang Sheet 03.');
  }

  let op03 = FS03_OPMAINT_find_(h03, [
    'Chi phí vận hành thuê trước VAT',
    'Chi phí vận hành trước VAT'
  ]);
  let maint03 = FS03_OPMAINT_find_(h03, ['Chi phí bảo trì trước VAT']);

  let nextCol = sh03.getLastColumn() + 1;
  if (op03 < 0) {
    sh03.getRange(1, nextCol).setValue('Chi phí vận hành thuê trước VAT');
    op03 = nextCol - 1;
    nextCol++;
  }
  if (maint03 < 0) {
    sh03.getRange(1, nextCol).setValue('Chi phí bảo trì trước VAT');
    maint03 = nextCol - 1;
  }

  h03 = sh03.getRange(1, 1, 1, sh03.getLastColumn()).getDisplayValues()[0];

  const n02 = Math.max(0, sh02.getLastRow() - 1);
  const byMonth = {};
  if (n02 > 0) {
    const d02 = sh02.getRange(2, 1, n02, sh02.getLastColumn()).getValues();
    d02.forEach(r => {
      const m = Number(r[month02]) || 0;
      if (!m) return;
      if (!byMonth[m]) byMonth[m] = { op: 0, maint: 0 };

      if (op02 >= 0 || maint02 >= 0) {
        byMonth[m].op += op02 >= 0 ? (Number(r[op02]) || 0) : 0;
        byMonth[m].maint += maint02 >= 0 ? (Number(r[maint02]) || 0) : 0;
      } else {
        // Cột gộp chỉ dùng làm phương án dự phòng. Không tự chia tùy tiện:
        // toàn bộ được ghi vào vận hành, bảo trì = 0 để tổng giá vốn không bị mất/đếm trùng.
        byMonth[m].op += Number(r[combined02]) || 0;
      }
    });
  }

  const n03 = Math.max(0, sh03.getLastRow() - 1);
  if (n03 > 0) {
    const months03 = sh03.getRange(2, month03 + 1, n03, 1).getValues();
    const opOut = [];
    const maintOut = [];
    months03.forEach(r => {
      const x = byMonth[Number(r[0]) || 0] || { op: 0, maint: 0 };
      opOut.push([x.op]);
      maintOut.push([x.maint]);
    });
    sh03.getRange(2, op03 + 1, n03, 1).setValues(opOut).setNumberFormat('#,##0');
    sh03.getRange(2, maint03 + 1, n03, 1).setValues(maintOut).setNumberFormat('#,##0');
  }

  sh03.getRange(1, op03 + 1).setFontWeight('bold').setWrap(true).setBackground('#D9EAD3');
  sh03.getRange(1, maint03 + 1).setFontWeight('bold').setWrap(true).setBackground('#D9EAD3');
  SpreadsheetApp.flush();
}

function FS03_OPMAINT_find_(headers, aliases) {
  const normalized = headers.map(FS03_OPMAINT_key_);
  for (const alias of aliases) {
    const idx = normalized.indexOf(FS03_OPMAINT_key_(alias));
    if (idx >= 0) return idx;
  }
  return -1;
}

function FS03_OPMAINT_key_(v) {
  return String(v || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}
