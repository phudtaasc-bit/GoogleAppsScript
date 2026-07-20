const FS03_CFG = Object.freeze({
  TECH: '01A. Kỹ thuật',
  REVENUE: '02. Doanh thu',
  COST: '03. Chi phí & Vốn',
  DEFAULT_VAT_RATE: 0.08
});

function FS_lapSheet03() {
  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName(FS03_CFG.TECH);
  const revenueSheet = ss.getSheetByName(FS03_CFG.REVENUE);

  if (!tech || !revenueSheet) {
    throw new Error('Cần lập "01A. Kỹ thuật" và "02. Doanh thu" trước.');
  }

  const months = Math.max(0, FS03_num_(FS03_readInfoValue_(tech, 'Số tháng mô hình')));
  const annualCostEscalation = FS03_rate_(
    FS03_readInfoValue_(tech, 'Tỷ lệ trượt chi phí/năm')
  );

  if (!months) throw new Error('Số tháng mô hình phải lớn hơn 0.');
  if (annualCostEscalation <= -1) {
    throw new Error('Tỷ lệ trượt chi phí/năm phải lớn hơn -100%.');
  }

  const products = FS03_readProducts_(tech);
  const productByCode = {};
  products.forEach(product => { productByCode[product.code] = product; });

  const totalArea = products.reduce((sum, product) => sum + product.area, 0);
  if (totalArea <= 0) throw new Error('Tổng DTKD của danh mục sản phẩm phải lớn hơn 0.');

  const costMap = FS03_readCostMap_(tech);
  const costItems = FS03_buildCostItems_(costMap);
  const schedules = FS03_readSchedules_(tech);
  const revenueRows = FS03_readRevenueRows_(revenueSheet, productByCode);

  FS03_validateLandCostStructure_(costMap);

  const scheduledByItemMonth = FS03_indexScheduledCosts_(costItems, schedules, months);
  const revenueByCodeMonth = FS03_indexRevenue_(revenueRows);
  const rows = [];

  for (let monthNo = 1; monthNo <= months; monthNo++) {
    const elapsedYears = (monthNo - 1) / 12;
    const costEscalationFactor = Math.pow(1 + annualCostEscalation, elapsedYears);

    products.forEach(product => {
      const key = product.code + '|' + monthNo;
      const revenue = revenueByCodeMonth[key] || FS03_emptyRevenue_(monthNo, product);
      const areaShare = product.area / totalArea;

      const constructionBase = FS03_scheduled_(scheduledByItemMonth, 'construction', monthNo) * areaShare;
      const infrastructureBase = FS03_scheduled_(scheduledByItemMonth, 'infrastructure', monthNo) * areaShare;
      const construction = constructionBase * costEscalationFactor;
      const infrastructure = infrastructureBase * costEscalationFactor;
      const clearance = FS03_scheduled_(scheduledByItemMonth, 'clearance', monthNo) * areaShare;
      const landUse = product.code === 'LK'
        ? FS03_scheduled_(scheduledByItemMonth, 'landUseLK', monthNo)
        : 0;
      const landRent = FS03_landRentForProduct_(product.code, scheduledByItemMonth, monthNo);
      const selling = product.group === 'Bán'
        ? revenue.saleRevenue * costItems.selling.rate
        : 0;
      const operating = product.group === 'Cho thuê'
        ? revenue.rentRevenue * product.operatingRate
        : 0;
      const maintenance = product.group === 'Cho thuê'
        ? revenue.rentRevenue * product.maintenanceRate
        : 0;
      const contingency = (construction + infrastructure) * costItems.contingency.rate;

      const totalBeforeVat =
        construction + clearance + infrastructure + landUse + landRent +
        selling + operating + maintenance + contingency;

      const vatInput =
        construction * costItems.construction.vatRate +
        clearance * costItems.clearance.vatRate +
        infrastructure * costItems.infrastructure.vatRate +
        landUse * costItems.landUseLK.vatRate +
        landRent * FS03_landRentVatRate_(product.code, costItems) +
        selling * costItems.selling.vatRate +
        operating * costItems.operating.vatRate +
        maintenance * costItems.maintenance.vatRate +
        contingency * costItems.contingency.vatRate;

      rows.push([
        monthNo,
        revenue.date,
        revenue.year,
        revenue.quarter,
        product.code,
        product.name,
        product.group,
        product.area,
        revenue.customerCash,
        revenue.vatOut,
        construction,
        clearance,
        infrastructure,
        landUse,
        landRent,
        selling,
        operating,
        maintenance,
        contingency,
        totalBeforeVat,
        vatInput,
        totalBeforeVat + vatInput
      ]);
    });
  }

  let sheet = ss.getSheetByName(FS03_CFG.COST);
  if (!sheet) sheet = ss.insertSheet(FS03_CFG.COST);
  sheet.clear();
  sheet.clearFormats();

  const headers = [[
    'Tháng số', 'Tháng', 'Năm', 'Quý', 'Mã SP', 'Tên sản phẩm', 'Nhóm', 'DTKD (m²)',
    'Dòng tiền khách hàng', 'VAT đầu ra',
    'XD/TB trước VAT', 'GPMB trước VAT', 'HTKT trước VAT',
    'Tiền SDĐ trước VAT', 'Tiền thuê đất trước VAT',
    'Chi phí bán hàng trước VAT', 'Chi phí vận hành trước VAT',
    'Chi phí bảo trì trước VAT', 'Chi phí dự phòng trước VAT',
    'Tổng chi trước VAT', 'VAT đầu vào', 'Tổng chi sau VAT'
  ]];

  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
  if (rows.length) sheet.getRange(2, 1, rows.length, headers[0].length).setValues(rows);
  FS03_format_(sheet, rows.length + 1, headers[0].length);
}

function FS03_readProducts_(sheet) {
  const rows = FS03_readBlock_(sheet, 'SAN_PHAM');
  const products = rows.map(row => ({
    code: String(row[0] || '').trim().toUpperCase(),
    name: String(row[1] || '').trim(),
    group: FS03_group_(row[2]),
    area: FS03_num_(row[3]),
    operatingRate: FS03_rate_(row[10]),
    maintenanceRate: FS03_rate_(row[11])
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

  return products;
}

function FS03_readCostMap_(sheet) {
  const rows = FS03_readBlock_(sheet, 'CHI_PHI_CHUNG');
  const map = {};

  rows.forEach(row => {
    const name = String(row[0] || '').trim();
    if (!name) return;
    map[FS03_key_(name)] = {
      name,
      beforeVat: FS03_num_(row[1]),
      vatRate: FS03_rate_(row[2]),
      afterVat: FS03_num_(row[3]),
      note: String(row[4] || '').trim(),
      rate: FS03_rate_(row[5])
    };
  });

  return map;
}

function FS03_buildCostItems_(costMap) {
  return {
    construction: FS03_cost_(costMap, ['Chi phí XD/TB/khác', 'Chi phí xây dựng & thiết bị', 'Chi phí XD/TB']),
    clearance: FS03_cost_(costMap, ['Chi phí GPMB']),
    infrastructure: FS03_cost_(costMap, ['Chi phí HTKT']),
    landUseLK: FS03_cost_(costMap, ['Tiền SDĐ Liền kề', 'Tiền SDD Liền kề']),
    landRentCC: FS03_cost_(costMap, ['Tiền thuê đất Chung cư']),
    landRentTMDV: FS03_cost_(costMap, ['Tiền thuê đất TMDV']),
    landRentCHO: FS03_cost_(costMap, ['Tiền thuê đất Chợ']),
    selling: FS03_cost_(costMap, ['Chi phí bán hàng']),
    operating: FS03_cost_(costMap, ['Chi phí vận hành']),
    maintenance: FS03_cost_(costMap, ['Chi phí bảo trì']),
    contingency: FS03_cost_(costMap, ['Chi phí dự phòng'])
  };
}

function FS03_cost_(costMap, aliases) {
  for (const alias of aliases) {
    const item = costMap[FS03_key_(alias)];
    if (item) return item;
  }
  return {
    name: aliases[0],
    beforeVat: 0,
    vatRate: FS03_CFG.DEFAULT_VAT_RATE,
    afterVat: 0,
    note: 'Mặc định VAT đầu vào 8%',
    rate: 0
  };
}

function FS03_readSchedules_(sheet) {
  return FS03_readBlock_(sheet, 'TIEN_DO_CHI_PHI')
    .map(row => ({
      item: String(row[0] || '').trim(),
      start: Math.max(1, FS03_num_(row[1])),
      duration: Math.max(1, FS03_num_(row[2])),
      rate: FS03_rate_(row[3]),
      type: String(row[4] || '').trim()
    }))
    .filter(schedule => schedule.item && schedule.rate !== 0);
}

function FS03_readRevenueRows_(sheet, productByCode) {
  if (sheet.getLastRow() <= 1) return [];

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const index = {};
  headers.forEach((header, position) => { index[FS03_key_(header)] = position; });

  const required = [
    'thangso', 'thang', 'nam', 'quy', 'masp',
    'doanhthubantruocvat', 'doanhthuthuetruocvat',
    'vatdaura', 'dongtienkhachhang'
  ];
  const missing = required.filter(key => index[key] == null);
  if (missing.length) throw new Error('Sheet 02 thiếu cột bắt buộc: ' + missing.join(', '));

  return values.map(row => {
    const code = String(row[index.masp] || '').trim().toUpperCase();
    if (!productByCode[code]) throw new Error('Sheet 02 có Mã SP không tồn tại trong SAN_PHAM: ' + code);
    return {
      monthNo: FS03_num_(row[index.thangso]),
      date: row[index.thang],
      year: FS03_num_(row[index.nam]),
      quarter: row[index.quy],
      code,
      saleRevenue: FS03_num_(row[index.doanhthubantruocvat]),
      rentRevenue: FS03_num_(row[index.doanhthuthuetruocvat]),
      vatOut: FS03_num_(row[index.vatdaura]),
      customerCash: FS03_num_(row[index.dongtienkhachhang])
    };
  });
}

function FS03_validateLandCostStructure_(costMap) {
  const legacyKeys = ['Tiền SDĐ/Thuê đất', 'Tiền SDĐ/thuê đất', 'Tiền sử dụng đất/thuê đất'];
  const legacy = legacyKeys.some(name => costMap[FS03_key_(name)] && costMap[FS03_key_(name)].beforeVat !== 0);
  if (legacy) throw new Error('CHI_PHI_CHUNG còn khoản mục tiền đất gộp. Phải tách theo từng sản phẩm.');
}

function FS03_indexScheduledCosts_(items, schedules, months) {
  const index = {};
  Object.keys(items).forEach(itemKey => {
    const item = items[itemKey];
    if (!item.beforeVat) return;
    const matched = schedules.filter(schedule => FS03_key_(schedule.item) === FS03_key_(item.name));

    matched.forEach(schedule => {
      const isOneTime = FS03_norm_(schedule.type).includes('mot lan');
      const end = Math.min(months, schedule.start + schedule.duration - 1);
      if (isOneTime) {
        const key = itemKey + '|' + schedule.start;
        index[key] = (index[key] || 0) + item.beforeVat * schedule.rate;
        return;
      }
      for (let monthNo = schedule.start; monthNo <= end; monthNo++) {
        const key = itemKey + '|' + monthNo;
        index[key] = (index[key] || 0) + item.beforeVat * schedule.rate / schedule.duration;
      }
    });
  });
  return index;
}

function FS03_indexRevenue_(rows) {
  const index = {};
  rows.forEach(row => {
    const key = row.code + '|' + row.monthNo;
    if (index[key]) throw new Error('Sheet 02 bị trùng khóa Mã SP + Tháng số: ' + key);
    index[key] = row;
  });
  return index;
}

function FS03_scheduled_(index, itemKey, monthNo) {
  return index[itemKey + '|' + monthNo] || 0;
}

function FS03_landRentForProduct_(code, index, monthNo) {
  if (code === 'CC') return FS03_scheduled_(index, 'landRentCC', monthNo);
  if (code === 'TMDV') return FS03_scheduled_(index, 'landRentTMDV', monthNo);
  if (code === 'CHO') return FS03_scheduled_(index, 'landRentCHO', monthNo);
  return 0;
}

function FS03_landRentVatRate_(code, items) {
  if (code === 'CC') return items.landRentCC.vatRate;
  if (code === 'TMDV') return items.landRentTMDV.vatRate;
  if (code === 'CHO') return items.landRentCHO.vatRate;
  return 0;
}

function FS03_emptyRevenue_(monthNo, product) {
  return {
    monthNo,
    date: '',
    year: '',
    quarter: '',
    code: product.code,
    saleRevenue: 0,
    rentRevenue: 0,
    vatOut: 0,
    customerCash: 0
  };
}

function FS03_readInfoValue_(sheet, label) {
  const target = FS03_key_(label);
  const values = sheet.getDataRange().getValues();
  for (let row = 0; row < values.length; row++) {
    if (FS03_key_(values[row][0]) === target) return values[row][1];
  }
  return '';
}

function FS03_readBlock_(sheet, marker) {
  const values = sheet.getDataRange().getValues();
  const target = FS03_key_(marker);
  let markerRow = -1;

  for (let row = 0; row < values.length; row++) {
    if (FS03_key_(values[row][0]) === target) {
      markerRow = row;
      break;
    }
  }

  if (markerRow < 0) throw new Error('Không tìm thấy block ' + marker + '.');
  const headerRow = markerRow + 1;
  const rows = [];
  let blanks = 0;

  for (let row = headerRow + 1; row < values.length; row++) {
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

function FS03_group_(value) {
  const key = FS03_key_(value);
  if (key === 'ban') return 'Bán';
  if (key === 'chothue') return 'Cho thuê';
  return String(value || '').trim();
}

function FS03_num_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  const text = String(value == null ? '' : value).trim().replace(/\s/g, '');
  if (!text) return 0;
  const normalized = text.includes(',') && text.includes('.')
    ? text.replace(/\./g, '').replace(',', '.')
    : text.replace(/,/g, '');
  const number = Number(normalized);
  return isFinite(number) ? number : 0;
}

function FS03_rate_(value) {
  if (typeof value === 'number') return value > 1 ? value / 100 : value;
  const text = String(value == null ? '' : value).trim();
  if (!text) return 0;
  const number = FS03_num_(text.replace('%', ''));
  return text.includes('%') || number > 1 ? number / 100 : number;
}

function FS03_norm_(value) {
  return String(value == null ? '' : value)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/\s+/g, ' ')
    .trim();
}

function FS03_key_(value) {
  return FS03_norm_(value)
    .replace(/²/g, '2')
    .replace(/\^2/g, '2')
    .replace(/m\s*2/g, 'm2')
    .replace(/[^a-z0-9]/g, '');
}

function FS03_format_(sheet, endRow, columnCount) {
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(7);
  sheet.getRange(1, 1, 1, columnCount)
    .setFontWeight('bold')
    .setBackground('#fce4d6')
    .setWrap(true)
    .setVerticalAlignment('middle');

  if (endRow > 1) {
    const rowCount = endRow - 1;
    sheet.getRange(2, 2, rowCount, 1).setNumberFormat('MM/yyyy');
    sheet.getRange(2, 8, rowCount, 15).setNumberFormat('#,##0');
  }

  const widths = [70, 85, 65, 90, 70, 180, 90, 95, 145, 105, 125, 115, 115, 125, 135, 145, 150, 145, 145, 135, 115, 135];
  widths.forEach((width, index) => sheet.setColumnWidth(index + 1, width));
  sheet.setRowHeight(1, 44);
}
