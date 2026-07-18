const FS00I_CFG = Object.freeze({
  SUMMARY: '00. Tổng hợp',
  TECH: '01A. Kỹ thuật',
  REVENUE: '02. Doanh thu',
  COST: '03. Chi phí & Vốn',
  TAX: '03A. Lợi nhuận & Thuế'
});

/**
 * Bản chạy ổn định cho phân tích hiệu quả từng sản phẩm.
 *
 * Khác với hàm cũ:
 * - Không phụ thuộc tuyệt đối vào ô "Số tháng mô hình" trên 01A.
 * - Nếu không đọc được chỉ tiêu này, số tháng được lấy theo Tháng số lớn nhất
 *   đang có trong 02, 03 và 03A.
 * - Không sửa dữ liệu các sheet lõi.
 */
function FS00I_phanTichHieuQuaTheoSanPham() {
  const ss = SpreadsheetApp.getActive();
  const summary = FS00F_sheet_(ss, FS00I_CFG.SUMMARY);
  const tech = FS00F_sheet_(ss, FS00I_CFG.TECH);
  const revenue = FS00F_readTable_(FS00F_sheet_(ss, FS00I_CFG.REVENUE));
  const cost = FS00F_readTable_(FS00F_sheet_(ss, FS00I_CFG.COST));
  const tax = FS00F_readTable_(FS00F_sheet_(ss, FS00I_CFG.TAX));

  FS00F_validateRevenue_(revenue);
  FS00F_validateCost_(cost);
  FS00F_validateTax_(tax);

  const monthsFromTech = Math.max(
    0,
    Math.round(FS00F_num_(FS00F_readInfo_(tech, 'Số tháng mô hình')))
  );
  const monthsFromData = Math.max(
    FS00I_maxMonth_(revenue),
    FS00I_maxMonth_(cost),
    FS00I_maxMonth_(tax)
  );
  const months = Math.max(monthsFromTech, monthsFromData);

  if (!months) {
    throw new Error(
      'Không xác định được số tháng mô hình từ 01A, 02, 03 hoặc 03A.'
    );
  }

  const loanRatio = FS00F_rate_(FS00F_readInfo_(tech, 'Tỷ lệ vốn vay'));
  if (loanRatio < 0 || loanRatio > 1) {
    throw new Error('Tỷ lệ vốn vay không hợp lệ.');
  }

  const products = FS00F_products_(revenue);
  if (!products.length) throw new Error('Không có sản phẩm để phân tích.');

  const rates = FS00F_discountRates_(summary);
  const layout = FS00C_chuanBiCotSanPham_(summary, products);

  const results = products.map(product => FS00F_analyzeProduct_({
    product,
    months,
    loanRatio,
    rates,
    revenue,
    cost,
    tax
  }));

  FS00F_writeSummary_(summary, layout, products, results);
  summary.setFrozenRows(0);
  summary.setFrozenColumns(0);
  SpreadsheetApp.flush();

  ss.toast(
    'Đã cập nhật phân tích hiệu quả theo từng sản phẩm.',
    'FS sản phẩm',
    6
  );
  return results;
}

function FS00I_maxMonth_(table) {
  if (!table || !table.rows || !table.alias || table.alias.monthNo == null) {
    return 0;
  }

  let maxMonth = 0;
  table.rows.forEach(row => {
    const monthNo = Math.round(FS00F_num_(row[table.alias.monthNo]));
    if (monthNo > maxMonth) maxMonth = monthNo;
  });
  return maxMonth;
}
