/*************************************************
 * ZZZZZZZZZZZZ_ConvergenceCombinedOpMaintFix.js
 * Vá trực tiếp hàm tính lại giá vốn khi Sheet 02 dùng:
 * - hai cột riêng Chi phí vận hành / Chi phí bảo trì; hoặc
 * - một cột gộp Chi phí vận hành & bảo trì.
 *************************************************/

FS_CONV_rebuildTaxCostIncludingOpMaint_ = function() {
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

  // Ưu tiên hai cột riêng. Chỉ dùng cột gộp khi không có đủ hai cột riêng.
  if (opCol > 0 || maintCol > 0) {
    if (opCol > 0) componentCols.push(opCol);
    if (maintCol > 0) componentCols.push(maintCol);
  } else if (combinedOpMaintCol > 0) {
    componentCols.push(combinedOpMaintCol);
  } else {
    throw new Error(
      'Sheet 02 không có cột chi phí vận hành/bảo trì. ' +
      'Các tiêu đề hiện có: ' + headers.filter(String).join(' | ')
    );
  }

  const uniqueComponentCols = [...new Set(componentCols)];
  if (!uniqueComponentCols.length) {
    throw new Error('Không tìm thấy các cột cấu phần giá vốn tại Sheet 02.');
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

    const totalTaxCost = uniqueComponentCols.reduce(
      (sum, c) => sum + (Number(r[c - 1]) || 0),
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
};
