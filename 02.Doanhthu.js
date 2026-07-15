function FS_lapSheet02() {
  const ss = SpreadsheetApp.getActive();
  const tech = FS_getSheet_(ss, FS_CFG.SHEETS.TECH, FS_CFG.SHEETS.TECH_LEGACY);
  if (!tech) throw new Error('Thiếu sheet "01A. Kỹ thuật".');

  const info = FS_readInfo_(tech);
  const months = FS_num_(info['Số tháng mô hình']);
  const start = info['Ngày bắt đầu dự án'];
  if (!months || !start) throw new Error('Thiếu thời gian mô hình.');

  let products = FS_readBlock_(tech, 'SAN_PHAM').map(r => ({
    name: String(r[0] || '').trim(),
    method: String(r[1] || '').trim(),
    area: FS_num_(r[2]),
    salePrice: FS_num_(r[3]),
    rentPrice: FS_num_(r[4]),
    vatRate: FS_rate_(r[6]),
    citRate: FS_rate_(r[7]),
    occupancy: FS_rate_(r[8]),
    rentUnit: String(r[12] || '').trim() || 'đ/m²/tháng'
  })).filter(x => x.name);
  products = FS_ganMaSanPham_(products);

  const plans = FS_readBlock_(tech, 'KE_HOACH_BAN_THU_TIEN').map(r => ({
    group: String(r[0] || ''),
    product: String(r[1] || ''),
    start: FS_num_(r[4]),
    duration: Math.max(1, FS_num_(r[5])),
    rate: FS_rate_(r[6])
  }));

  const growth = FS_rate_(info['Tỷ lệ tăng giá/năm']);
  const rows = [];

  for (let t = 1; t <= months; t++) {
    const date = FS_addMonths_(start, t - 1);
    const factor = Math.pow(1 + growth, (t - 1) / 12);

    for (const p of products) {
      const loaiHinh = FS_loaiHinh_(p.method);
      const applicable = plans.filter(x =>
        FS_key_(x.product) === FS_key_(p.name) &&
        FS_norm_(x.group).includes('thu tien') &&
        t >= x.start && t < x.start + x.duration
      );

      const progress = loaiHinh === 'Bán'
        ? applicable.reduce((s, x) => s + x.rate / x.duration, 0)
        : applicable.reduce((s, x) => s + x.rate, 0);

      const salePrice = p.salePrice * factor;
      const rentPrice = p.rentPrice * factor;
      const saleRevenue = loaiHinh === 'Bán' ? p.area * salePrice * progress : 0;
      const rentRevenue = loaiHinh === 'Cho thuê' ? p.area * rentPrice * p.occupancy * progress : 0;
      const revenue = saleRevenue + rentRevenue;
      const vatOut = revenue * p.vatRate;

      rows.push([
        t, date, date.getFullYear(), 'Q' + Math.ceil((date.getMonth() + 1) / 3) + '/' + date.getFullYear(),
        p.code, p.name, loaiHinh, p.area, salePrice, rentPrice, p.rentUnit,
        p.occupancy, progress, saleRevenue, rentRevenue, revenue,
        p.vatRate, vatOut, revenue + vatOut, p.citRate
      ]);
    }
  }

  const sh = FS_getOrCreateSheet_(ss, FS_CFG.SHEETS.REVENUE);
  FS_resetSheet_(sh, rows.length + 1, 20);
  sh.getRange(1, 1, 1, 20).setValues([[
    'Tháng số', 'Tháng', 'Năm', 'Quý', 'Mã SP', 'Tên sản phẩm', 'Loại hình',
    'DTKD (m²)', 'Giá bán (đ/m²)', 'Đơn giá thuê', 'ĐVT đơn giá thuê',
    'Tỷ lệ lấp đầy', 'Tiến độ ghi nhận doanh thu (%)',
    'Doanh thu bán trước VAT', 'Doanh thu thuê trước VAT', 'Tổng doanh thu trước VAT',
    'Thuế suất VAT (%)', 'VAT đầu ra', 'Dòng tiền thu khách hàng', 'Thuế suất TNDN (%)'
  ]]);
  if (rows.length) sh.getRange(2, 1, rows.length, 20).setValues(rows);
  FS02_format_(sh, rows.length + 1);
}

function FS02_format_(sh, endRow) {
  sh.setFrozenRows(1);
  sh.setFrozenColumns(7);
  sh.getRange(1, 1, 1, 20).setFontWeight('bold').setBackground('#d9eaf7').setWrap(true);
  if (endRow > 1) {
    sh.getRange(2, 2, endRow - 1, 1).setNumberFormat('MM/yyyy');
    sh.getRange(2, 8, endRow - 1, 1).setNumberFormat('#,##0.00');
    sh.getRange(2, 9, endRow - 1, 2).setNumberFormat('#,##0.00');
    sh.getRange(2, 12, endRow - 1, 2).setNumberFormat('0.00%');
    sh.getRange(2, 14, endRow - 1, 6).setNumberFormat('#,##0.00');
    sh.getRange(2, 17, endRow - 1, 1).setNumberFormat('0.00%');
    sh.getRange(2, 20, endRow - 1, 1).setNumberFormat('0.00%');
  }
  sh.autoResizeColumns(1, 20);
}
