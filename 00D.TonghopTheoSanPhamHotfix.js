/**
 * Hotfix entry point cho phân tích hiệu quả theo sản phẩm.
 *
 * Khác biệt so với hàm gốc:
 * - Không bắt buộc Sheet 03A phải có cột "Lãi vay phân bổ".
 * - Nếu không có cột này, lãi vay sản phẩm được đọc là 0.
 * - Không sửa các module lõi và không chạy lại hội tụ lãi vay.
 */
function FS00B_phanTichHieuQuaTheoSanPham_V2() {
  const ss = SpreadsheetApp.getActive();
  const sheets = FS00B_getRequiredSheets_(ss);

  const revenue = FS00B_readTable_(sheets.revenue);
  const cost = FS00B_readTable_(sheets.cost);
  const tax = FS00B_readTable_(sheets.tax);
  const cash = FS00B_readTable_(sheets.cash);

  FS00B_require_(revenue, [
    'thangso', 'thang', 'masp', 'tensanpham', 'nhom',
    'tongdoanhthutruocvat', 'vatdaura', 'dongtienkhachhang'
  ], sheets.revenue.getName());

  FS00B_require_(cost, [
    'thangso', 'masp', 'tongchitruocvat', 'vatdauvao', 'tongchisauvat',
    'xdtbtruocvat', 'gpmbtruocvat', 'tiensddtruocvat',
    'tienthuedattruocvat', 'htkttruocvat',
    'chiphibanhangtruocvat', 'chiphivanhanhtruocvat',
    'chiphibaotritruocvat', 'chiphiduphongtruocvat'
  ], sheets.cost.getName());

  FS00B_require_(tax, [
    'thangso', 'masp', 'thuetndn', 'lnst'
  ], sheets.tax.getName());

  FS00B_require_(cash, [
    'thangso', 'thang', 'giainganvay', 'tragoc'
  ], sheets.cash.getName());

  const products = FS00B_getProducts_(revenue);
  if (!products.length) {
    throw new Error('Không tìm thấy sản phẩm trong Sheet 02.');
  }

  const layout = FS00C_chuanBiCotSanPham_(sheets.summary, products);
  const projectRates = FS00B_readDiscountRates_(sheets.summary);

  const projectFinance = FS00B_buildProjectFinance_(cash);
  const allProductCosts = FS00B_buildAllProductCostMap_(cost, products);
  const totalCostByMonth = FS00B_sumCostByMonth_(allProductCosts);
  const globalCostShares = FS00B_globalCostShares_(allProductCosts, products);

  const results = products.map(product => FS00B_analyzeProduct_({
    product,
    revenue,
    cost,
    tax,
    projectFinance,
    totalCostByMonth,
    globalCostShare: globalCostShares[product.code] || 0,
    projectRates
  }));

  FS00B_writeResults_(sheets.summary, layout, products, results);
  SpreadsheetApp.flush();

  ss.toast(
    'Đã cập nhật NPV, IRR và các chỉ tiêu hiệu quả cho ' + products.length + ' sản phẩm.',
    'Phân tích theo sản phẩm',
    6
  );

  return results;
}
