function FS_lapSheet00_TheoDanhMuc() {
  const result = FS_lapSheet00();
  FS00_capNhatNhanSanPham_();
  return result;
}

function FS00_capNhatNhanSanPham_() {
  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName('01A. Kỹ thuật');
  const summary = ss.getSheetByName('00. Tổng hợp');
  if (!tech || !summary) return;

  const block = FS00_findBlock_(tech, 'SAN_PHAM');
  const values = tech.getRange(block.startRow, 1, block.rowCount, 14).getValues();
  const products = [];
  const seen = {};

  values.forEach(row => {
    const code = String(row[0] || '').trim().toUpperCase();
    const name = String(row[1] || '').trim();
    const group = String(row[2] || '').trim();
    if (!code || !name || seen[code]) return;
    seen[code] = true;
    products.push({ code, name, group });
  });

  const totalRevenueRow = FS00_findSummaryRow_(summary, 'Tổng doanh thu có VAT');
  const totalCostRow = FS00_findSummaryRow_(summary, 'Tổng chi phí có VAT');
  if (!totalRevenueRow || !totalCostRow) return;

  const detailCount = Math.min(products.length, totalCostRow - totalRevenueRow - 1);
  for (let i = 0; i < detailCount; i++) {
    const product = products[i];
    const label = 'Phần ' + product.name + (product.group ? ' - ' + product.group : '');
    summary.getRange(totalRevenueRow + 1 + i, 2).setValue(label);
  }
}
