/*************************************************
 * ZZZZZZZZZZZZ_OpMaintCombinedBranchFix.js
 * Sửa trực tiếp lỗi nhận diện cột vận hành/bảo trì tại Sheet 02.
 * Hỗ trợ cả:
 * - 2 cột riêng: Chi phí vận hành..., Chi phí bảo trì...
 * - 1 cột gộp: Chi phí vận hành & bảo trì thuê trước VAT
 *************************************************/

function FS_CONV_rebuildTaxCostIncludingOpMaint_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('02. Doanh thu');
  if (!sh) throw new Error('Không tìm thấy Sheet 02. Doanh thu.');

  const headerRow = FS_CONV_detectHeaderRow_(sh, [
    'Tổng giá vốn tính thuế',
    'Lợi nhuận chịu thuế',
    'Thuế TNDN tạm tính'
  ]);

  const headers = sh.getRange(headerRow, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  const col = names => FS_CONV_findCol_(headers, names);
  const required = names => {
    const c = col(names);
    if (c < 1) throw new Error('Sheet 02 thiếu cột: ' + names.join(' / '));
    return c;
  };

  const monthCol = required(['Tháng số']);
  const revenueCol = required(['Tổng doanh thu trước VAT']);
  const taxRateCol = required(['Thuế TNDN %']);
  const totalTaxCostCol = required(['Tổng giá vốn tính thuế']);
  const taxableProfitCol = required(['Lợi nhuận chịu thuế']);
  const citCol = required(['Thuế TNDN tạm tính']);

  const opCol = col([
    'Chi phí vận hành thuê trước VAT',
    'Chi phí vận hành trước VAT'
  ]);

  const maintCol = col([
    'Chi phí bảo trì trước VAT'
  ]);

  const combinedOpMaintCol = col([
    'Chi phí vận hành & bảo trì thuê trước VAT',
    'Chi phí vận hành và bảo trì thuê trước VAT',
    'Chi phí vận hành & bảo trì trước VAT',
    'Chi phí vận hành và bảo trì trước VAT'
  ]);

  const componentCols = [
    col(['CP XD/TB trực tiếp trước VAT']),
    col(['Chi phí bán hàng trước VAT']),
    col(['Chi phí GPMB phân bổ trước VAT']),
    col(['Chi phí HTKT phân bổ trước VAT']),
    col(['Tiền SDĐ/thuê đất phân bổ trước VAT']),
    col(['Chi phí dự phòng phân bổ trước VAT']),
    col(['Chi phí lãi vay phân bổ'])
  ].filter(c => c > 0);

  if (combinedOpMaintCol > 0) {
    componentCols.push(combinedOpMaintCol);
  } else {
    if (opCol > 0) componentCols.push(opCol);
    if (maintCol > 0) componentCols.push(maintCol);
  }

  const uniqueCols = [...new Set(componentCols)];
  if (!uniqueCols.length) {
    throw new Error('Không tìm thấy các cột cấu phần giá vốn tại Sheet 02.');
  }

  if (combinedOpMaintCol < 1 && opCol < 1 && maintCol < 1) {
    throw new Error(
      'Không tìm thấy cột vận hành/bảo trì tại Sheet 02. Header hiện có: ' +
      headers.filter(v => String(v || '').trim() !== '').join(' | ')
    );
  }

  const startRow = headerRow + 1;
  const lastRow = sh.getLastRow();
  if (lastRow < startRow) return;

  const n = lastRow - startRow + 1;
  const width = sh.getLastColumn();
  const data = sh.getRange(startRow, 1, n, width).getValues();

  const outCost = [];
  const outProfit = [];
  const outCit = [];

  data.forEach(r => {
    const month = Number(r[monthCol - 1]) || 0;
    if (!month) {
      outCost.push(['']);
      outProfit.push(['']);
      outCit.push(['']);
      return;
    }

    const totalTaxCost = uniqueCols.reduce(
      (s, c) => s + (Number(r[c - 1]) || 0),
      0
    );

    const revenue = Number(r[revenueCol - 1]) || 0;
    const taxRate = FS_CONV_rate_(r[taxRateCol - 1]);
    const taxableProfit = Math.max(0, revenue - totalTaxCost);
    const cit = taxableProfit * taxRate;

    outCost.push([totalTaxCost]);
    outProfit.push([taxableProfit]);
    outCit.push([cit]);
  });

  sh.getRange(startRow, totalTaxCostCol, n, 1).setValues(outCost);
  sh.getRange(startRow, taxableProfitCol, n, 1).setValues(outProfit);
  sh.getRange(startRow, citCol, n, 1).setValues(outCit);
  SpreadsheetApp.flush();
}
