function FS_lapSheet02() {
  const ss = SpreadsheetApp.getActive();
  const tech = FS_getSheet_(ss, FS_CFG.SHEETS.TECH, FS_CFG.SHEETS.TECH_LEGACY);
  if (!tech) throw new Error('Thiếu sheet "01A. Kỹ thuật".');

  const info = FS_readInfo_(tech);
  const months = Math.max(0, FS_num_(info['Số tháng mô hình']));
  const startDate = info['Ngày bắt đầu dự án'];

  if (!months) throw new Error('Số tháng mô hình phải lớn hơn 0.');
  if (!(startDate instanceof Date) || isNaN(startDate.getTime())) {
    throw new Error('Ngày bắt đầu dự án không hợp lệ.');
  }

  const products = FS_readBlock_(tech, 'SAN_PHAM')
    .map(row => ({
      code: String(row[0] || '').trim().toUpperCase(),
      name: String(row[1] || '').trim(),
      group: FS02_group_(row[2]),
      area: FS_num_(row[3]),
      salePrice: FS_num_(row[4]),
      rentPrice: FS_num_(row[5]),
      vatRate: FS_rate_(row[7]),
      occupancy: FS_rate_(row[9])
    }))
    .filter(product => product.code || product.name);

  FS02_validateProducts_(products);

  const codeByProductName = {};
  products.forEach(product => {
    const nameKey = FS_key_(product.name);
    if (nameKey) codeByProductName[nameKey] = product.code;
  });

  const plans = FS_readBlock_(tech, 'KE_HOACH_BAN_THU_TIEN')
    .map(row => {
      const productName = String(row[1] || '').trim();
      return {
        code: codeByProductName[FS_key_(productName)] || '',
        productName,
        start: Math.max(1, FS_num_(row[4])),
        duration: Math.max(1, FS_num_(row[5])),
        rate: FS_rate_(row[6])
      };
    })
    .filter(plan => plan.productName || plan.code);

  const unresolvedPlans = plans.filter(plan => !plan.code);
  if (unresolvedPlans.length) {
    const names = [...new Set(unresolvedPlans.map(plan => plan.productName).filter(Boolean))];
    throw new Error(
      'Kế hoạch bán/thu tiền có sản phẩm chưa ánh xạ được Mã SP: ' + names.join(', ')
    );
  }

  const plansByCodeMonth = FS02_indexPlans_(plans, months);
  const annualGrowth = FS_rate_(info['Tỷ lệ tăng giá/năm']);
  const rows = [];

  for (let monthNo = 1; monthNo <= months; monthNo++) {
    const date = FS_addMonths_(startDate, monthNo - 1);
    const priceFactor = Math.pow(1 + annualGrowth, (monthNo - 1) / 12);

    products.forEach(product => {
      const planKey = product.code + '|' + monthNo;
      const activePlans = plansByCodeMonth[planKey] || [];

      const progress = product.group === 'Bán'
        ? activePlans.reduce((sum, plan) => sum + plan.rate / plan.duration, 0)
        : activePlans.reduce((sum, plan) => sum + plan.rate, 0);

      const salePrice = product.salePrice * priceFactor;
      const rentPrice = product.rentPrice * priceFactor;

      const saleRevenue = product.group === 'Bán'
        ? product.area * salePrice * progress
        : 0;

      const rentRevenue = product.group === 'Cho thuê'
        ? product.area * rentPrice * product.occupancy * progress
        : 0;

      const totalRevenue = saleRevenue + rentRevenue;
      const vatOut = totalRevenue * product.vatRate;
      const customerCash = totalRevenue + vatOut;

      rows.push([
        monthNo,
        date,
        date.getFullYear(),
        'Q' + Math.ceil((date.getMonth() + 1) / 3) + '/' + date.getFullYear(),
        product.code,
        product.name,
        product.group,
        product.area,
        salePrice,
        rentPrice,
        product.occupancy,
        progress,
        saleRevenue,
        rentRevenue,
        totalRevenue,
        product.vatRate,
        vatOut,
        customerCash
      ]);
    });
  }

  const sheet = FS_getOrCreateSheet_(ss, FS_CFG.SHEETS.REVENUE);
  FS_resetSheet_(sheet, rows.length + 1, 18);

  sheet.getRange(1, 1, 1, 18).setValues([[
    'Tháng số',
    'Tháng',
    'Năm',
    'Quý',
    'Mã SP',
    'Tên sản phẩm',
    'Nhóm',
    'DTKD (m²)',
    'Giá bán trước thuế/m²',
    'Giá thuê/m²/tháng',
    'Tỷ lệ lấp đầy',
    'Tiến độ doanh thu',
    'Doanh thu bán trước VAT',
    'Doanh thu thuê trước VAT',
    'Tổng doanh thu trước VAT',
    'Thuế suất VAT',
    'VAT đầu ra',
    'Dòng tiền khách hàng'
  ]]);

  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 18).setValues(rows);
  }

  FS02_format_(sheet, rows.length + 1);
}

function FS02_group_(value) {
  const key = FS_key_(value);
  if (key === 'ban') return 'Bán';
  if (key === 'chothue') return 'Cho thuê';
  return String(value || '').trim();
}

function FS02_validateProducts_(products) {
  if (!products.length) throw new Error('Block SAN_PHAM không có dữ liệu sản phẩm.');

  const missingCodes = products.filter(product => !product.code);
  if (missingCodes.length) throw new Error('Block SAN_PHAM còn thiếu Mã SP.');

  const codes = products.map(product => product.code);
  const duplicatedCodes = [...new Set(codes.filter((code, index) => codes.indexOf(code) !== index))];
  if (duplicatedCodes.length) {
    throw new Error('Mã SP bị trùng trong block SAN_PHAM: ' + duplicatedCodes.join(', '));
  }

  const invalidGroups = products
    .filter(product => product.group !== 'Bán' && product.group !== 'Cho thuê')
    .map(product => product.code + ': ' + product.group);

  if (invalidGroups.length) {
    throw new Error('Nhóm sản phẩm chỉ được là "Bán" hoặc "Cho thuê": ' + invalidGroups.join('; '));
  }
}

function FS02_indexPlans_(plans, months) {
  const index = {};

  plans.forEach(plan => {
    const endMonth = Math.min(months, plan.start + plan.duration - 1);
    for (let monthNo = plan.start; monthNo <= endMonth; monthNo++) {
      const key = plan.code + '|' + monthNo;
      if (!index[key]) index[key] = [];
      index[key].push(plan);
    }
  });

  return index;
}

function FS02_format_(sheet, endRow) {
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(7);

  sheet.getRange(1, 1, 1, 18)
    .setFontWeight('bold')
    .setBackground('#d9eaf7')
    .setWrap(true)
    .setVerticalAlignment('middle');

  if (endRow > 1) {
    const rowCount = endRow - 1;
    sheet.getRange(2, 2, rowCount, 1).setNumberFormat('MM/yyyy');
    sheet.getRange(2, 8, rowCount, 3).setNumberFormat('#,##0');
    sheet.getRange(2, 11, rowCount, 2).setNumberFormat('0.00%');
    sheet.getRange(2, 13, rowCount, 3).setNumberFormat('#,##0');
    sheet.getRange(2, 16, rowCount, 1).setNumberFormat('0.00%');
    sheet.getRange(2, 17, rowCount, 2).setNumberFormat('#,##0');
  }

  const widths = [70, 85, 65, 90, 70, 180, 90, 95, 145, 145, 105, 115, 150, 150, 155, 105, 120, 150];
  widths.forEach((width, index) => sheet.setColumnWidth(index + 1, width));
  sheet.setRowHeight(1, 42);
}
