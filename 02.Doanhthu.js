const FS02_CFG = Object.freeze({
  TECH: '01A. Kỹ thuật',
  REVENUE: '02. Doanh thu'
});

function FS_lapSheet02() {
  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName(FS02_CFG.TECH);
  if (!tech) throw new Error('Thiếu sheet "01A. Kỹ thuật".');

  const months = Math.max(0, FS02_num_(FS02_readInfoValue_(tech, 'Số tháng mô hình')));
  const startDate = FS02_readInfoValue_(tech, 'Ngày bắt đầu dự án');
  const annualGrowth = FS02_rate_(FS02_readInfoValue_(tech, 'Tỷ lệ tăng giá/năm'));

  if (!months) throw new Error('Số tháng mô hình phải lớn hơn 0.');
  if (!(startDate instanceof Date) || isNaN(startDate.getTime())) {
    throw new Error('Ngày bắt đầu dự án không hợp lệ.');
  }

  const productRows = FS02_readBlock_(tech, 'SAN_PHAM');
  const products = productRows.map(row => ({
    code: String(row[0] || '').trim().toUpperCase(),
    name: String(row[1] || '').trim(),
    group: FS02_group_(row[2]),
    area: FS02_num_(row[3]),
    salePrice: FS02_num_(row[4]),
    rentPrice: FS02_num_(row[5]),
    vatRate: FS02_rate_(row[7]),
    occupancy: FS02_rate_(row[9])
  })).filter(product => product.code || product.name);

  FS02_validateProducts_(products);

  const codeByProductName = {};
  products.forEach(product => {
    const key = FS02_key_(product.name);
    if (key) codeByProductName[key] = product.code;
  });

  const planRows = FS02_readBlock_(tech, 'KE_HOACH_BAN_THU_TIEN');
  const plans = planRows.map(row => {
    const productName = String(row[1] || '').trim();
    return {
      code: codeByProductName[FS02_key_(productName)] || '',
      productName,
      start: Math.max(1, FS02_num_(row[4])),
      duration: Math.max(1, FS02_num_(row[5])),
      rate: FS02_rate_(row[6])
    };
  }).filter(plan => plan.productName || plan.code);

  const unresolvedPlans = plans.filter(plan => !plan.code);
  if (unresolvedPlans.length) {
    const names = [...new Set(unresolvedPlans.map(plan => plan.productName).filter(Boolean))];
    throw new Error('Kế hoạch bán/thu tiền có sản phẩm chưa ánh xạ được Mã SP: ' + names.join(', '));
  }

  const plansByCodeMonth = FS02_indexPlans_(plans, months);
  const rows = [];

  for (let monthNo = 1; monthNo <= months; monthNo++) {
    const date = FS02_addMonths_(startDate, monthNo - 1);
    const priceFactor = Math.pow(1 + annualGrowth, (monthNo - 1) / 12);

    products.forEach(product => {
      const activePlans = plansByCodeMonth[product.code + '|' + monthNo] || [];
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
        totalRevenue + vatOut
      ]);
    });
  }

  let sheet = ss.getSheetByName(FS02_CFG.REVENUE);
  if (!sheet) sheet = ss.insertSheet(FS02_CFG.REVENUE);
  sheet.clear();
  sheet.clearFormats();

  sheet.getRange(1, 1, 1, 18).setValues([[
    'Tháng số', 'Tháng', 'Năm', 'Quý', 'Mã SP', 'Tên sản phẩm', 'Nhóm',
    'DTKD (m²)', 'Giá bán trước thuế/m²', 'Giá thuê/m²/tháng',
    'Tỷ lệ lấp đầy', 'Tiến độ doanh thu', 'Doanh thu bán trước VAT',
    'Doanh thu thuê trước VAT', 'Tổng doanh thu trước VAT',
    'Thuế suất VAT', 'VAT đầu ra', 'Dòng tiền khách hàng'
  ]]);

  if (rows.length) sheet.getRange(2, 1, rows.length, 18).setValues(rows);
  FS02_format_(sheet, rows.length + 1);
}

function FS02_readInfoValue_(sheet, label) {
  const target = FS02_key_(label);
  const values = sheet.getDataRange().getValues();
  for (let r = 0; r < values.length; r++) {
    if (FS02_key_(values[r][0]) === target) return values[r][1];
  }
  return '';
}

function FS02_readBlock_(sheet, marker) {
  const values = sheet.getDataRange().getValues();
  const markerKey = FS02_key_(marker);
  let markerRow = -1;

  for (let r = 0; r < values.length; r++) {
    if (FS02_key_(values[r][0]) === markerKey) {
      markerRow = r;
      break;
    }
  }

  if (markerRow < 0) throw new Error('Không tìm thấy block ' + marker + '.');
  const headerRow = markerRow + 1;
  if (headerRow >= values.length) throw new Error('Block ' + marker + ' không có hàng tiêu đề.');

  const rows = [];
  let blankCount = 0;

  for (let r = headerRow + 1; r < values.length; r++) {
    const row = values[r];
    const first = String(row[0] || '').trim();
    const hasData = row.some(value => String(value == null ? '' : value).trim() !== '');
    const isNextBlock = /^[A-Z0-9_]+$/.test(first) && first.indexOf('_') >= 0;

    if (isNextBlock) break;
    if (!hasData) {
      blankCount++;
      if (blankCount >= 2) break;
      continue;
    }

    blankCount = 0;
    rows.push(row);
  }

  return rows;
}

function FS02_group_(value) {
  const key = FS02_key_(value);
  if (key === 'ban') return 'Bán';
  if (key === 'chothue') return 'Cho thuê';
  return String(value || '').trim();
}

function FS02_validateProducts_(products) {
  if (!products.length) throw new Error('Block SAN_PHAM không có dữ liệu sản phẩm.');
  if (products.some(product => !product.code)) throw new Error('Block SAN_PHAM còn thiếu Mã SP.');

  const codes = products.map(product => product.code);
  const duplicated = [...new Set(codes.filter((code, index) => codes.indexOf(code) !== index))];
  if (duplicated.length) throw new Error('Mã SP bị trùng trong block SAN_PHAM: ' + duplicated.join(', '));

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

function FS02_addMonths_(date, months) {
  return new Date(date.getFullYear(), date.getMonth() + months, 1);
}

function FS02_num_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  const text = String(value == null ? '' : value).trim().replace(/\s/g, '');
  if (!text) return 0;
  const normalized = text.indexOf(',') >= 0 && text.indexOf('.') >= 0
    ? text.replace(/\./g, '').replace(',', '.')
    : text.replace(/,/g, '');
  const number = Number(normalized);
  return isFinite(number) ? number : 0;
}

function FS02_rate_(value) {
  if (typeof value === 'number') return value > 1 ? value / 100 : value;
  const text = String(value == null ? '' : value).trim();
  if (!text) return 0;
  const number = FS02_num_(text.replace('%', ''));
  return text.indexOf('%') >= 0 || number > 1 ? number / 100 : number;
}

function FS02_norm_(value) {
  return String(value == null ? '' : value)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/\s+/g, ' ')
    .trim();
}

function FS02_key_(value) {
  return FS02_norm_(value)
    .replace(/²/g, '2')
    .replace(/\^2/g, '2')
    .replace(/m\s*2/g, 'm2')
    .replace(/[^a-z0-9]/g, '');
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