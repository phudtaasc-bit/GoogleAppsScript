const FS03A_CFG = Object.freeze({
  TECH: '01A. Kỹ thuật',
  REVENUE: '02. Doanh thu',
  COST: '03. Chi phí & Vốn',
  PROFIT: '03A. Lợi nhuận & Thuế',
  RECONCILIATION_TOLERANCE: 1
});

function FS_lapSheet03A() {
  return FS_hoiTuTaiTro();
}

function FS03A_build_(interestByMonth, writeSheet) {
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
  const capitalizedInterestByCode = FS03A_allocateInterest_(costRows, interestByMonth || [], products);
  const pools = FS03A_buildPools_(products, revenueRows, costRows, capitalizedInterestByCode);
  const recognitionRates = FS03A_buildRecognitionRates_(products, revenueRows, pools);
  const rows = [];

  revenueRows.forEach(revenue => {
    const product = productByCode[revenue.code];
    const key = revenue.code + '|' + revenue.monthNo;
    const cost = costByKey[key] || FS03A_emptyCost_(revenue);
    const pool = pools[revenue.code];
    const recognitionRate = recognitionRates[key] || 0;

    const baseCostRecognized = pool.baseCostPool * recognitionRate;
    const capitalizedInterestRecognized = pool.interestPool * recognitionRate;
    const landUseRecognized = pool.landUsePool * recognitionRate;
    const landRentRecognized = pool.landRentPool * recognitionRate;

    const selling = cost.selling;
    const operating = cost.operating;
    const maintenance = cost.maintenance;

    const totalCost =
      baseCostRecognized +
      capitalizedInterestRecognized +
      landUseRecognized +
      landRentRecognized +
      selling +
      operating +
      maintenance;

    const pbt = revenue.totalRevenue - totalCost;
    const taxable = Math.max(0, pbt);
    const cit = taxable * revenue.citRate;
    const pat = pbt - cit;

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
      capitalizedInterestRecognized,
      landUseRecognized,
      landRentRecognized,
      selling,
      operating,
      maintenance,
      totalCost,
      pbt,
      taxable,
      revenue.citRate,
      cit,
      pat
    ]);
  });

  FS03A_validateInterestAllocation_(rows, interestByMonth || []);

  if (writeSheet !== false) FS03A_write_(ss, rows);
  return rows;
}

function FS03A_allocateInterest_(costRows, interestByMonth, products) {
  const monthTotals = {};
  const eligibleByMonthAndCode = {};
  const productCodes = products.map(product => product.code);

  costRows.forEach(row => {
    const eligible = FS03A_eligibleInvestment_(row);
    monthTotals[row.monthNo] = (monthTotals[row.monthNo] || 0) + eligible;
    eligibleByMonthAndCode[row.monthNo + '|' + row.code] = eligible;
  });

  const result = {};
  productCodes.forEach(code => { result[code] = 0; });

  interestByMonth.forEach((value, index) => {
    const monthNo = index + 1;
    const interest = FS03A_num_(value);
    if (!interest) return;

    const totalEligible = monthTotals[monthNo] || 0;
    if (totalEligible > 0) {
      let allocatedSoFar = 0;
      const activeCodes = productCodes.filter(code =>
        FS03A_num_(eligibleByMonthAndCode[monthNo + '|' + code]) > 0
      );

      activeCodes.forEach((code, position) => {
        const eligible = FS03A_num_(eligibleByMonthAndCode[monthNo + '|' + code]);
        const allocated = position === activeCodes.length - 1
          ? interest - allocatedSoFar
          : interest * eligible / totalEligible;

        result[code] += allocated;
        allocatedSoFar += allocated;
      });
    } else {
      // Khi tháng phát sinh lãi vay nhưng không còn chi phí đầu tư đủ điều kiện,
      // phân bổ theo tỷ trọng pool đầu tư lũy kế của sản phẩm; nếu toàn bộ bằng 0,
      // dồn vào sản phẩm đầu tiên để bảo đảm không làm thất thoát lãi vay.
      const cumulativeEligible = {};
      let totalCumulative = 0;

      productCodes.forEach(code => {
        const amount = costRows
          .filter(row => row.code === code && row.monthNo <= monthNo)
          .reduce((sum, row) => sum + FS03A_eligibleInvestment_(row), 0);
        cumulativeEligible[code] = amount;
        totalCumulative += amount;
      });

      if (totalCumulative > 0) {
        let allocatedSoFar = 0;
        const activeCodes = productCodes.filter(code => cumulativeEligible[code] > 0);
        activeCodes.forEach((code, position) => {
          const allocated = position === activeCodes.length - 1
            ? interest - allocatedSoFar
            : interest * cumulativeEligible[code] / totalCumulative;
          result[code] += allocated;
          allocatedSoFar += allocated;
        });
      } else if (productCodes.length) {
        result[productCodes[0]] += interest;
      }
    }
  });

  return result;
}

function FS03A_eligibleInvestment_(row) {
  return row.construction +
    row.clearance +
    row.infrastructure +
    row.contingency +
    row.landUse +
    row.landRent;
}

function FS03A_buildPools_(products, revenueRows, costRows, interestByCode) {
  const pools = {};

  products.forEach(product => {
    pools[product.code] = {
      totalSaleRevenue: 0,
      firstRentMonth: 0,
      lastModelMonth: 0,
      baseCostPool: 0,
      interestPool: FS03A_num_(interestByCode[product.code]),
      landUsePool: 0,
      landRentPool: 0
    };
  });

  revenueRows.forEach(row => {
    const pool = pools[row.code];
    pool.totalSaleRevenue += row.saleRevenue;
    pool.lastModelMonth = Math.max(pool.lastModelMonth, row.monthNo);
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

function FS03A_buildRecognitionRates_(products, revenueRows, pools) {
  const rowsByCode = {};
  products.forEach(product => { rowsByCode[product.code] = []; });
  revenueRows.forEach(row => rowsByCode[row.code].push(row));

  const rates = {};

  products.forEach(product => {
    const rows = (rowsByCode[product.code] || []).slice().sort((a, b) => a.monthNo - b.monthNo);
    if (!rows.length) return;

    if (product.group === 'Bán') {
      const saleRows = rows.filter(row => row.saleRevenue > 0);

      if (saleRows.length && pools[product.code].totalSaleRevenue > 0) {
        let accumulatedRate = 0;
        saleRows.forEach((row, position) => {
          const rate = position === saleRows.length - 1
            ? Math.max(0, 1 - accumulatedRate)
            : row.saleRevenue / pools[product.code].totalSaleRevenue;
          rates[row.code + '|' + row.monthNo] = rate;
          accumulatedRate += rate;
        });
      } else {
        const finalRow = rows[rows.length - 1];
        rates[finalRow.code + '|' + finalRow.monthNo] = 1;
      }
      return;
    }

    const firstRentMonth = pools[product.code].firstRentMonth;
    const recognitionRows = firstRentMonth > 0
      ? rows.filter(row => row.monthNo >= firstRentMonth)
      : [rows[rows.length - 1]];

    let accumulatedRate = 0;
    recognitionRows.forEach((row, position) => {
      const rate = position === recognitionRows.length - 1
        ? Math.max(0, 1 - accumulatedRate)
        : 1 / recognitionRows.length;
      rates[row.code + '|' + row.monthNo] = rate;
      accumulatedRate += rate;
    });
  });

  return rates;
}

function FS03A_validateInterestAllocation_(rows, interestByMonth) {
  const totalInputInterest = interestByMonth.reduce(
    (sum, value) => sum + FS03A_num_(value),
    0
  );
  const totalRecognizedInterest = rows.reduce(
    (sum, row) => sum + FS03A_num_(row[11]),
    0
  );
  const difference = totalRecognizedInterest - totalInputInterest;

  if (Math.abs(difference) > FS03A_CFG.RECONCILIATION_TOLERANCE) {
    throw new Error(
      'Lãi vay vốn hóa phân bổ tại Sheet 03A không khớp lãi vay đầu vào. ' +
      'Đầu vào: ' + Math.round(totalInputInterest).toLocaleString('vi-VN') +
      '; đã phân bổ: ' + Math.round(totalRecognizedInterest).toLocaleString('vi-VN') +
      '; chênh lệch: ' + Math.round(difference).toLocaleString('vi-VN') + ' đồng.'
    );
  }
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
  if (invalid.length) throw new Error('Nhóm sản phẩm không hợp lệ: ' + invalid.map(product => product.code).join(', '));

  const missingLease = products.filter(product => product.group === 'Cho thuê' && product.leaseYears <= 0);
  if (missingLease.length) {
    throw new Error('Sản phẩm cho thuê chưa có Thời gian thuê: ' + missingLease.map(product => product.code).join(', '));
  }

  return products;
}

function FS03A_readRevenue_(sheet, productByCode) {
  const table = FS03A_readTable_(sheet);
  const required = [
    'thangso', 'thang', 'nam', 'quy', 'masp',
    'doanhthubantruocvat', 'doanhthuthuetruocvat',
    'tongdoanhthutruocvat', 'thuesuattndn'
  ];
  FS03A_require_(table.index, required, '02. Doanh thu');

  return table.values.map(row => {
    const code = String(row[table.index.masp] || '').trim().toUpperCase();
    if (!productByCode[code]) throw new Error('Sheet 02 có Mã SP không tồn tại: ' + code);

    return {
      monthNo: FS03A_num_(row[table.index.thangso]),
      date: row[table.index.thang],
      year: row[table.index.nam],
      quarter: row[table.index.quy],
      code,
      saleRevenue: FS03A_num_(row[table.index.doanhthubantruocvat]),
      rentRevenue: FS03A_num_(row[table.index.doanhthuthuetruocvat]),
      totalRevenue: FS03A_num_(row[table.index.tongdoanhthutruocvat]),
      citRate: FS03A_rate_(row[table.index.thuesuattndn])
    };
  });
}

function FS03A_readCosts_(sheet, productByCode) {
  const table = FS03A_readTable_(sheet);
  const required = [
    'thangso', 'thang', 'nam', 'quy', 'masp',
    'xdtbtruocvat', 'gpmbtruocvat', 'htkttruocvat',
    'tiensddtruocvat', 'tienthuedattruocvat',
    'chiphiduphongtruocvat', 'chiphibanhangtruocvat',
    'chiphivanhanhtruocvat', 'chiphibaotritruocvat'
  ];
  FS03A_require_(table.index, required, '03. Chi phí & Vốn');

  return table.values.map(row => {
    const code = String(row[table.index.masp] || '').trim().toUpperCase();
    if (!productByCode[code]) throw new Error('Sheet 03 có Mã SP không tồn tại: ' + code);

    return {
      monthNo: FS03A_num_(row[table.index.thangso]),
      date: row[table.index.thang],
      year: row[table.index.nam],
      quarter: row[table.index.quy],
      code,
      construction: FS03A_num_(row[table.index.xdtbtruocvat]),
      clearance: FS03A_num_(row[table.index.gpmbtruocvat]),
      infrastructure: FS03A_num_(row[table.index.htkttruocvat]),
      landUse: FS03A_num_(row[table.index.tiensddtruocvat]),
      landRent: FS03A_num_(row[table.index.tienthuedattruocvat]),
      contingency: FS03A_num_(row[table.index.chiphiduphongtruocvat]),
      selling: FS03A_num_(row[table.index.chiphibanhangtruocvat]),
      operating: FS03A_num_(row[table.index.chiphivanhanhtruocvat]),
      maintenance: FS03A_num_(row[table.index.chiphibaotritruocvat])
    };
  });
}

function FS03A_indexByKey_(rows, source) {
  const result = {};
  rows.forEach(row => {
    const key = row.code + '|' + row.monthNo;
    if (result[key]) throw new Error(source + ' bị trùng khóa Mã SP + Tháng số: ' + key);
    result[key] = row;
  });
  return result;
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

function FS03A_write_(ss, rows) {
  let sheet = ss.getSheetByName(FS03A_CFG.PROFIT);
  if (!sheet) sheet = ss.insertSheet(FS03A_CFG.PROFIT);

  sheet.clear();
  sheet.clearFormats();

  const headers = [[
    'Tháng số', 'Tháng', 'Năm', 'Quý', 'Mã SP', 'Tên sản phẩm', 'Nhóm',
    'Doanh thu bán trước VAT', 'Doanh thu thuê trước VAT', 'Tổng doanh thu trước VAT',
    'Giá vốn XD/GPMB/HTKT/Dự phòng', 'Lãi vay vốn hóa phân bổ',
    'Tiền SDĐ phân bổ', 'Tiền thuê đất phân bổ',
    'Chi phí bán hàng', 'Chi phí vận hành', 'Chi phí bảo trì',
    'Tổng chi phí hạch toán', 'Lợi nhuận trước thuế', 'Thu nhập chịu thuế',
    'Thuế suất TNDN', 'Thuế TNDN', 'LNST'
  ]];

  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
  if (rows.length) sheet.getRange(2, 1, rows.length, headers[0].length).setValues(rows);

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(7);
  sheet.getRange(1, 1, 1, headers[0].length)
    .setFontWeight('bold')
    .setBackground('#e2f0d9')
    .setWrap(true);

  if (rows.length) {
    sheet.getRange(2, 2, rows.length, 1).setNumberFormat('MM/yyyy');
    sheet.getRange(2, 8, rows.length, 13).setNumberFormat('#,##0');
    sheet.getRange(2, 21, rows.length, 1).setNumberFormat('0.00%');
    sheet.getRange(2, 22, rows.length, 2).setNumberFormat('#,##0');
  }

  sheet.autoResizeColumns(1, headers[0].length);
}

function FS03A_readTable_(sheet) {
  const columnCount = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, columnCount).getDisplayValues()[0];
  const values = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, columnCount).getValues()
    : [];
  const index = {};
  headers.forEach((header, position) => { index[FS03A_key_(header)] = position; });
  return { values, index };
}

function FS03A_require_(index, required, sheetName) {
  const missing = required.filter(key => index[key] == null);
  if (missing.length) throw new Error('Sheet "' + sheetName + '" thiếu cột: ' + missing.join(', '));
}

function FS03A_readBlock_(sheet, marker) {
  const values = sheet.getDataRange().getValues();
  const target = FS03A_key_(marker);
  let start = -1;

  for (let row = 0; row < values.length; row++) {
    if (FS03A_key_(values[row][0]) === target) {
      start = row;
      break;
    }
  }

  if (start < 0) throw new Error('Không tìm thấy block ' + marker + '.');

  const rows = [];
  let blankCount = 0;

  for (let row = start + 2; row < values.length; row++) {
    const current = values[row];
    const first = String(current[0] || '').trim();
    const hasData = current.some(value => String(value == null ? '' : value).trim() !== '');

    if (/^[A-Z0-9_]+$/.test(first) && first.includes('_')) break;
    if (!hasData) {
      if (++blankCount >= 2) break;
      continue;
    }

    blankCount = 0;
    rows.push(current);
  }

  return rows;
}

function FS03A_group_(value) {
  const key = FS03A_key_(value);
  return key === 'ban' ? 'Bán' : key === 'chothue' ? 'Cho thuê' : String(value || '').trim();
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
