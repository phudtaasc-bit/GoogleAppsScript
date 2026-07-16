const FS03A_CFG = Object.freeze({
  TECH: '01A. Kỹ thuật',
  REVENUE: '02. Doanh thu',
  COST: '03. Chi phí & Vốn',
  PROFIT: '03A. Lợi nhuận & Thuế'
});

function FS_lapSheet03A() {
  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName(FS03A_CFG.TECH);
  const revenueSheet = ss.getSheetByName(FS03A_CFG.REVENUE);
  const costSheet = ss.getSheetByName(FS03A_CFG.COST);

  if (!tech || !revenueSheet || !costSheet) {
    throw new Error('Cần lập "01A. Kỹ thuật", "02. Doanh thu" và "03. Chi phí & Vốn" trước.');
  }

  const products = FS03A_readProducts_(tech);
  const productByCode = {};
  products.forEach(product => { productByCode[product.code] = product; });

  const revenueRows = FS03A_readRevenue_(revenueSheet, productByCode);
  const costRows = FS03A_readCosts_(costSheet, productByCode);
  const costByKey = FS03A_indexByKey_(costRows, 'Sheet 03');

  const totals = FS03A_buildPools_(products, revenueRows, costRows);
  const rows = [];

  revenueRows.forEach(revenue => {
    const product = productByCode[revenue.code];
    const key = revenue.code + '|' + revenue.monthNo;
    const cost = costByKey[key] || FS03A_emptyCost_(revenue);
    const pool = totals[revenue.code];

    let baseCostRecognized = 0;
    let landUseRecognized = 0;
    let landRentRecognized = 0;

    if (product.group === 'Bán') {
      const recognitionRate = pool.totalSaleRevenue > 0
        ? revenue.saleRevenue / pool.totalSaleRevenue
        : 0;

      baseCostRecognized = pool.baseCostPool * recognitionRate;
      landUseRecognized = pool.landUsePool * recognitionRate;
      landRentRecognized = pool.landRentPool * recognitionRate;
    } else {
      const leaseMonths = Math.max(1, product.leaseYears * 12);
      const active = pool.firstRentMonth > 0 &&
        revenue.monthNo >= pool.firstRentMonth &&
        revenue.monthNo < pool.firstRentMonth + leaseMonths;

      if (active) {
        baseCostRecognized = pool.baseCostPool / leaseMonths;
        landRentRecognized = pool.landRentPool / leaseMonths;
      }
    }

    const sellingExpense = cost.selling;
    const operatingExpense = cost.operating;
    const maintenanceExpense = cost.maintenance;

    const totalRecognizedCost =
      baseCostRecognized + landUseRecognized + landRentRecognized +
      sellingExpense + operatingExpense + maintenanceExpense;

    const profitBeforeTax = revenue.totalRevenue - totalRecognizedCost;
    const taxableIncome = Math.max(0, profitBeforeTax);
    const cit = taxableIncome * revenue.citRate;
    const profitAfterTax = profitBeforeTax - cit;

    rows.push([
      revenue.monthNo,
      revenue.date,
      revenue.year,
      revenue.quarter,
      revenue.code,
      product.name,
      product.group,
      revenue.saleRevenue,
      revenue.rentRevenue,
      revenue.totalRevenue,
      baseCostRecognized,
      landUseRecognized,
      landRentRecognized,
      sellingExpense,
      operatingExpense,
      maintenanceExpense,
      totalRecognizedCost,
      profitBeforeTax,
      taxableIncome,
      revenue.citRate,
      cit,
      profitAfterTax
    ]);
  });

  let sheet = ss.getSheetByName(FS03A_CFG.PROFIT);
  if (!sheet) sheet = ss.insertSheet(FS03A_CFG.PROFIT);
  sheet.clear();
  sheet.clearFormats();

  const headers = [[
    'Tháng số', 'Tháng', 'Năm', 'Quý', 'Mã SP', 'Tên sản phẩm', 'Nhóm',
    'Doanh thu bán trước VAT', 'Doanh thu thuê trước VAT', 'Tổng doanh thu trước VAT',
    'Giá vốn XD/GPMB/HTKT/Dự phòng', 'Tiền SDĐ phân bổ', 'Tiền thuê đất phân bổ',
    'Chi phí bán hàng', 'Chi phí vận hành', 'Chi phí bảo trì',
    'Tổng chi phí hạch toán', 'Lợi nhuận trước thuế', 'Thu nhập chịu thuế',
    'Thuế suất TNDN', 'Thuế TNDN', 'LNST'
  ]];

  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
  if (rows.length) sheet.getRange(2, 1, rows.length, headers[0].length).setValues(rows);
  FS03A_format_(sheet, rows.length + 1, headers[0].length);
}

function FS03A_readProducts_(sheet) {
  const rows = FS03A_readBlock_(sheet, 'SAN_PHAM');
  const products = rows.map(row => ({
    code: String(row[0] || '').trim().toUpperCase(),
    name: String(row[1] || '').trim(),
    group: FS03A_group_(row[2]),
    leaseYears: Math.max(0, FS03A_num_(row[12]))
  })).filter(product => product.code || product.name);

  if (!products.length) throw new Error('Block SAN_PHAM không có dữ liệu sản phẩm.');
  if (products.some(product => !product.code)) throw new Error('Block SAN_PHAM còn thiếu Mã SP.');

  const codes = products.map(product => product.code);
  const duplicated = [...new Set(codes.filter((code, index) => codes.indexOf(code) !== index))];
  if (duplicated.length) throw new Error('Mã SP bị trùng trong SAN_PHAM: ' + duplicated.join(', '));

  const invalid = products.filter(product => !['Bán', 'Cho thuê'].includes(product.group));
  if (invalid.length) {
    throw new Error('Nhóm sản phẩm không hợp lệ: ' + invalid.map(product => product.code + ': ' + product.group).join('; '));
  }

  const missingLease = products.filter(product => product.group === 'Cho thuê' && product.leaseYears <= 0);
  if (missingLease.length) {
    throw new Error('Sản phẩm cho thuê chưa có Thời gian thuê: ' + missingLease.map(product => product.code).join(', '));
  }

  return products;
}

function FS03A_readRevenue_(sheet, productByCode) {
  if (sheet.getLastRow() <= 1) return [];

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const index = FS03A_headerIndex_(headers);
  const required = [
    'thangso', 'thang', 'nam', 'quy', 'masp',
    'doanhthubantruocvat', 'doanhthuthuetruocvat', 'tongdoanhthutruocvat', 'thuesuattndn'
  ];
  const missing = required.filter(key => index[key] == null);
  if (missing.length) throw new Error('Sheet 02 thiếu cột bắt buộc: ' + missing.join(', '));

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  return values.map(row => {
    const code = String(row[index.masp] || '').trim().toUpperCase();
    if (!productByCode[code]) throw new Error('Sheet 02 có Mã SP không tồn tại trong SAN_PHAM: ' + code);

    return {
      monthNo: FS03A_num_(row[index.thangso]),
      date: row[index.thang],
      year: FS03A_num_(row[index.nam]),
      quarter: row[index.quy],
      code,
      saleRevenue: FS03A_num_(row[index.doanhthubantruocvat]),
      rentRevenue: FS03A_num_(row[index.doanhthuthuetruocvat]),
      totalRevenue: FS03A_num_(row[index.tongdoanhthutruocvat]),
      citRate: FS03A_rate_(row[index.thuesuattndn])
    };
  });
}

function FS03A_readCosts_(sheet, productByCode) {
  if (sheet.getLastRow() <= 1) return [];

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const index = FS03A_headerIndex_(headers);
  const required = [
    'thangso', 'thang', 'nam', 'quy', 'masp',
    'xdtbtruocvat', 'gpmbtruocvat', 'htkttruocvat',
    'tiensddtruocvat', 'tienthuedattruocvat',
    'chiphiduphongtruocvat', 'chiphibanhangtruocvat',
    'chiphivanhanhtruocvat', 'chiphibaotritruocvat'
  ];
  const missing = required.filter(key => index[key] == null);
  if (missing.length) throw new Error('Sheet 03 thiếu cột bắt buộc: ' + missing.join(', '));

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  return values.map(row => {
    const code = String(row[index.masp] || '').trim().toUpperCase();
    if (!productByCode[code]) throw new Error('Sheet 03 có Mã SP không tồn tại trong SAN_PHAM: ' + code);

    return {
      monthNo: FS03A_num_(row[index.thangso]),
      date: row[index.thang],
      year: FS03A_num_(row[index.nam]),
      quarter: row[index.quy],
      code,
      construction: FS03A_num_(row[index.xdtbtruocvat]),
      clearance: FS03A_num_(row[index.gpmbtruocvat]),
      infrastructure: FS03A_num_(row[index.htkttruocvat]),
      landUse: FS03A_num_(row[index.tiensddtruocvat]),
      landRent: FS03A_num_(row[index.tienthuedattruocvat]),
      contingency: FS03A_num_(row[index.chiphiduphongtruocvat]),
      selling: FS03A_num_(row[index.chiphibanhangtruocvat]),
      operating: FS03A_num_(row[index.chiphivanhanhtruocvat]),
      maintenance: FS03A_num_(row[index.chiphibaotritruocvat])
    };
  });
}

function FS03A_buildPools_(products, revenueRows, costRows) {
  const pools = {};
  products.forEach(product => {
    pools[product.code] = {
      totalSaleRevenue: 0,
      firstRentMonth: 0,
      baseCostPool: 0,
      landUsePool: 0,
      landRentPool: 0
    };
  });

  revenueRows.forEach(row => {
    const pool = pools[row.code];
    pool.totalSaleRevenue += row.saleRevenue;
    if (!pool.firstRentMonth && row.rentRevenue > 0) pool.firstRentMonth = row.monthNo;
  });

  costRows.forEach(row => {
    const pool = pools[row.code];
    pool.baseCostPool += row.construction + row.clearance + row.infrastructure + row.contingency;
    pool.landUsePool += row.landUse;
    pool.landRentPool += row.landRent;
  });

  return pools;
}

function FS03A_indexByKey_(rows, sourceName) {
  const index = {};
  rows.forEach(row => {
    const key = row.code + '|' + row.monthNo;
    if (index[key]) throw new Error(sourceName + ' bị trùng khóa Mã SP + Tháng số: ' + key);
    index[key] = row;
  });
  return index;
}

function FS03A_emptyCost_(revenue) {
  return {
    monthNo: revenue.monthNo,
    date: revenue.date,
    year: revenue.year,
    quarter: revenue.quarter,
    code: revenue.code,
    construction: 0,
    clearance: 0,
    infrastructure: 0,
    landUse: 0,
    landRent: 0,
    contingency: 0,
    selling: 0,
    operating: 0,
    maintenance: 0
  };
}

function FS03A_readBlock_(sheet, marker) {
  const values = sheet.getDataRange().getValues();
  const target = FS03A_key_(marker);
  let markerRow = -1;

  for (let row = 0; row < values.length; row++) {
    if (FS03A_key_(values[row][0]) === target) {
      markerRow = row;
      break;
    }
  }

  if (markerRow < 0) throw new Error('Không tìm thấy block ' + marker + '.');
  const rows = [];
  let blanks = 0;

  for (let row = markerRow + 2; row < values.length; row++) {
    const current = values[row];
    const first = String(current[0] || '').trim();
    const hasData = current.some(value => String(value == null ? '' : value).trim() !== '');
    const nextBlock = /^[A-Z0-9_]+$/.test(first) && first.includes('_');

    if (nextBlock) break;
    if (!hasData) {
      if (++blanks >= 2) break;
      continue;
    }

    blanks = 0;
    rows.push(current);
  }

  return rows;
}

function FS03A_headerIndex_(headers) {
  const index = {};
  headers.forEach((header, position) => { index[FS03A_key_(header)] = position; });
  return index;
}

function FS03A_group_(value) {
  const key = FS03A_key_(value);
  if (key === 'ban') return 'Bán';
  if (key === 'chothue') return 'Cho thuê';
  return String(value || '').trim();
}

function FS03A_num_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  const text = String(value == null ? '' : value).trim().replace(/\s/g, '');
  if (!text) return 0;
  const normalized = text.includes(',') && text.includes('.')
    ? text.replace(/\./g, '').replace(',', '.')
    : text.replace(/,/g, '');
  const number = Number(normalized);
  return isFinite(number) ? number : 0;
}

function FS03A_rate_(value) {
  if (typeof value === 'number') return value > 1 ? value / 100 : value;
  const text = String(value == null ? '' : value).trim();
  if (!text) return 0;
  const number = FS03A_num_(text.replace('%', ''));
  return text.includes('%') || number > 1 ? number / 100 : number;
}

function FS03A_norm_(value) {
  return String(value == null ? '' : value)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/\s+/g, ' ')
    .trim();
}

function FS03A_key_(value) {
  return FS03A_norm_(value)
    .replace(/²/g, '2')
    .replace(/\^2/g, '2')
    .replace(/m\s*2/g, 'm2')
    .replace(/[^a-z0-9]/g, '');
}

function FS03A_format_(sheet, endRow, columnCount) {
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(7);

  sheet.getRange(1, 1, 1, columnCount)
    .setFontWeight('bold')
    .setBackground('#e2f0d9')
    .setWrap(true)
    .setVerticalAlignment('middle');

  if (endRow > 1) {
    const rowCount = endRow - 1;
    sheet.getRange(2, 2, rowCount, 1).setNumberFormat('MM/yyyy');
    sheet.getRange(2, 8, rowCount, 12).setNumberFormat('#,##0');
    sheet.getRange(2, 20, rowCount, 1).setNumberFormat('0.00%');
    sheet.getRange(2, 21, rowCount, 2).setNumberFormat('#,##0');
  }

  const widths = [70, 85, 65, 90, 70, 180, 90, 145, 145, 155, 165, 130, 145, 135, 135, 125, 145, 145, 135, 110, 120, 125];
  widths.forEach((width, index) => sheet.setColumnWidth(index + 1, width));
  sheet.setRowHeight(1, 46);
}
