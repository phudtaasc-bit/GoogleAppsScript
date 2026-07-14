/*************************************************
 * ZZZZZZZZZZZZ_CombinedOpMaintHeaderFix.gs
 *
 * Vá tương thích Sheet 02 khi mô hình dùng cột gộp:
 * "Chi phí vận hành & bảo trì thuê trước VAT"
 * thay vì hai cột vận hành và bảo trì riêng.
 *************************************************/

const FS_COMB_BASE_LAP_SHEET02_ = FS_lapSheet02;

FS_lapSheet02 = function() {
  try {
    FS_COMB_BASE_LAP_SHEET02_();
  } catch (err) {
    const msg = String(err && err.message || err || '');
    const isOpMaintHeaderError =
      msg.includes('Chi phí vận hành thuê trước VAT') ||
      msg.includes('Chi phí vận hành trước VAT') ||
      msg.includes('Chi phí bảo trì trước VAT');

    if (!isOpMaintHeaderError) throw err;

    // Hàm gốc đã dựng Sheet 02 trước khi patch hậu xử lý báo thiếu cột.
    // Tiếp tục bằng bộ tính giá vốn tương thích cột gộp.
    FS_COMB_rebuildTaxCost_();
  }
};

function FS_COMB_rebuildTaxCost_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('02. Doanh thu');
  if (!sh) throw new Error('Không tìm thấy Sheet 02. Doanh thu.');

  const headerRow = FS_COMB_detectHeaderRow_(sh);
  const headers = sh.getRange(headerRow, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  const find = names => FS_COMB_findCol_(headers, names);
  const req = names => {
    const c = find(names);
    if (c < 1) throw new Error('Sheet 02 thiếu cột bắt buộc: ' + names.join(' / '));
    return c;
  };

  const monthCol = req(['Tháng số']);
  const revenueCol = req(['Tổng doanh thu trước VAT']);
  const taxRateCol = req(['Thuế TNDN %']);
  const totalCostCol = req(['Tổng giá vốn tính thuế']);
  const taxableCol = req(['Lợi nhuận chịu thuế']);
  const citCol = req(['Thuế TNDN tạm tính']);

  const opCol = find(['Chi phí vận hành thuê trước VAT', 'Chi phí vận hành trước VAT']);
  const maintCol = find(['Chi phí bảo trì trước VAT']);
  const combinedCol = find([
    'Chi phí vận hành & bảo trì thuê trước VAT',
    'Chi phí vận hành và bảo trì thuê trước VAT',
    'Chi phí vận hành & bảo trì trước VAT',
    'Chi phí vận hành và bảo trì trước VAT'
  ]);

  const componentCols = [
    find(['CP XD/TB trực tiếp trước VAT']),
    find(['Chi phí bán hàng trước VAT']),
    find(['Chi phí GPMB phân bổ trước VAT']),
    find(['Chi phí HTKT phân bổ trước VAT']),
    find(['Tiền SDĐ/thuê đất phân bổ trước VAT']),
    find(['Chi phí dự phòng phân bổ trước VAT']),
    find(['Chi phí lãi vay phân bổ'])
  ].filter(c => c > 0);

  // Nếu có hai cột riêng thì dùng hai cột riêng; nếu không thì dùng đúng một cột gộp.
  if (opCol > 0 || maintCol > 0) {
    if (opCol > 0) componentCols.push(opCol);
    if (maintCol > 0) componentCols.push(maintCol);
  } else if (combinedCol > 0) {
    componentCols.push(combinedCol);
  } else {
    throw new Error(
      'Sheet 02 không có cột chi phí vận hành/bảo trì riêng hoặc cột gộp vận hành & bảo trì.'
    );
  }

  const startRow = headerRow + 1;
  const lastRow = sh.getLastRow();
  if (lastRow < startRow) return;

  const n = lastRow - startRow + 1;
  const data = sh.getRange(startRow, 1, n, sh.getLastColumn()).getValues();
  const outCost = [];
  const outTaxable = [];
  const outCit = [];

  data.forEach(row => {
    const month = Number(row[monthCol - 1]) || 0;
    if (!month) {
      outCost.push(['']);
      outTaxable.push(['']);
      outCit.push(['']);
      return;
    }

    const totalCost = componentCols.reduce(
      (sum, c) => sum + (Number(row[c - 1]) || 0),
      0
    );
    const revenue = Number(row[revenueCol - 1]) || 0;
    const taxRate = FS_COMB_rate_(row[taxRateCol - 1]);
    const taxable = Math.max(0, revenue - totalCost);

    outCost.push([totalCost]);
    outTaxable.push([taxable]);
    outCit.push([taxable * taxRate]);
  });

  sh.getRange(startRow, totalCostCol, n, 1).setValues(outCost);
  sh.getRange(startRow, taxableCol, n, 1).setValues(outTaxable);
  sh.getRange(startRow, citCol, n, 1).setValues(outCit);
  SpreadsheetApp.flush();
}

function FS_COMB_detectHeaderRow_(sh) {
  const maxRows = Math.min(6, sh.getLastRow());
  for (let r = 1; r <= maxRows; r++) {
    const h = sh.getRange(r, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
    if (FS_COMB_findCol_(h, ['Tổng giá vốn tính thuế']) > 0 &&
        FS_COMB_findCol_(h, ['Thuế TNDN tạm tính']) > 0) return r;
  }
  throw new Error('Không nhận diện được dòng tiêu đề Sheet 02.');
}

function FS_COMB_findCol_(headers, names) {
  const keys = headers.map(FS_COMB_key_);
  for (const name of names) {
    const idx = keys.indexOf(FS_COMB_key_(name));
    if (idx >= 0) return idx + 1;
  }
  return -1;
}

function FS_COMB_rate_(v) {
  let n = Number(v);
  if (!isFinite(n)) return 0;
  if (Math.abs(n) > 1) n /= 100;
  return Math.max(0, n);
}

function FS_COMB_key_(v) {
  return String(v || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/&/g, ' va ')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}
