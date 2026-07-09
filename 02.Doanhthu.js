/*************************************************
 * 02_DoanhThu.gs - FS V2.0
 * Thay toàn bộ file / hàm lập Sheet 02 bằng bản này.
 *
 * Mục tiêu:
 * - Giữ nguyên layout 32 cột A:AF của Sheet 02.
 * - Sản phẩm bán: chi phí đầu tư + lãi vay phân bổ theo tiến độ bán/thu tiền.
 * - Sản phẩm cho thuê: không có chi phí bán hàng; CPXD, GPMB, HTKT,
 *   tiền thuê đất, dự phòng, lãi vay phân bổ đều trong thời gian thuê
 *   giống logic khấu hao/phân bổ nguyên giá.
 * - Tiền thuê đất Chung cư xử lý như tiền SDĐ của Liền kề: phân bổ theo tiến độ bán.
 *************************************************/

function FS_lapSheet02() {
  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName('01. Kỹ thuật');
  if (!tech) throw new Error('Không tìm thấy sheet "01. Kỹ thuật".');

  const sh = FS02V2_getSheetByPrefix_(ss, '02');
  if (!sh) throw new Error('Không tìm thấy sheet bắt đầu bằng "02".');

  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  if (!sh03) {
    throw new Error('Không tìm thấy sheet "03. Chi phí & Vốn". Cần chạy FS_lapSheet03() trước FS_lapSheet02().');
  }

  const info = FS02V2_readInfo_(tech);
  const soThang = Number(info['Số tháng mô hình']) || 0;
  const startDate = info['Ngày bắt đầu dự án'];

  if (!soThang || !startDate) {
    throw new Error('Thiếu "Số tháng mô hình" hoặc "Ngày bắt đầu dự án" tại sheet 01. Kỹ thuật.');
  }

  const products = FS02V2_readProducts_(tech);
  const schedules = FS02V2_readSchedules_(tech);
  const commonCosts = FS02V2_readCommonCosts_(tech);

  if (!products.length) throw new Error('Không có dữ liệu sản phẩm trong block SAN_PHAM.');

  const nSP = products.length;
  const totalRows = soThang * nSP;
  const endRow = totalRows + 1;
  const endRow03 = soThang + 1;

  FS02V2_prepareSheet_(sh, endRow, 32);

  const headers = [[
    'Tháng số',
    'Tháng',
    'Năm',
    'Quý',
    'Loại sản phẩm',
    'Hình thức',
    'DTKD',
    'Giá bán/m2',
    'Giá thuê/m2/tháng',
    'CPXD/m2',
    'Lấp đầy thuê',
    'CPVH thuê',
    'Diện tích đất',
    'Tiến độ thu',
    'Doanh thu bán trước VAT',
    'Doanh thu thuê trước VAT',
    'Tổng doanh thu trước VAT',
    'VAT đầu ra %',
    'VAT đầu ra',
    'Dòng tiền huy động từ KH',
    'Thuế TNDN %',
    'CP XD/TB trực tiếp trước VAT',
    'Chi phí bán hàng trước VAT',
    'Chi phí vận hành thuê trước VAT',
    'Chi phí GPMB phân bổ trước VAT',
    'Chi phí HTKT phân bổ trước VAT',
    'Tiền SDĐ/thuê đất phân bổ trước VAT',
    'Chi phí dự phòng phân bổ trước VAT',
    'Chi phí lãi vay phân bổ',
    'Tổng giá vốn tính thuế',
    'Lợi nhuận chịu thuế',
    'Thuế TNDN tạm tính'
  ]];
  sh.getRange(1, 1, 1, 32).setValues(headers);

  const totals03 = FS02V2_readTotals03_(sh03, endRow03);
  const totalCPXD = totals03.cpxd;
  const totalGPMB = totals03.gpmb;
  const totalHTKT = totals03.htkt;
  const totalDuPhong = totals03.duPhong;
  const totalLaiVay = totals03.laiVay;

  const saleCostRate = FS02V2_getCostRate_(commonCosts, ['Chi phí bán hàng']);

  const productMeta = FS02V2_buildProductMeta_(
    products,
    schedules,
    commonCosts,
    {
      totalCPXD,
      totalGPMB,
      totalHTKT,
      totalDuPhong,
      totalLaiVay
    }
  );

  const rows = [];
  const tz = Session.getScriptTimeZone();

  for (let t = 1; t <= soThang; t++) {
    const monthDate = FS02V2_addMonths_(startDate, t - 1);
    const year = monthDate.getFullYear();
    const quarter = 'Q' + Math.ceil((monthDate.getMonth() + 1) / 3) + '/' + year;

    products.forEach(p => {
      const key = FS02V2_key_(p.name);
      const meta = productMeta[key];

      const isSale = FS02V2_isSale_(p.method);
      const isRent = FS02V2_isRent_(p.method);

      const saleProgress = isSale ? FS02V2_getSaleProgress_(p.name, t, schedules) : 0;
      const rentRate = isRent ? FS02V2_getRentActiveRate_(p.name, t, schedules) : 0;

      const saleRevenue = isSale ? p.area * p.salePrice * saleProgress : 0;
      const rentRevenue = isRent ? p.area * p.rentPrice * p.occupancy * rentRate : 0;
      const revenue = saleRevenue + rentRevenue;

      const vatOut = revenue * p.vatOut;
      const cashFromCustomer = revenue + vatOut;

      const saleDenom = meta.totalSaleProgress || 1;
      const rentDuration = meta.rentDuration || 1;

      let cpxdAlloc = 0;
      let gpmbAlloc = 0;
      let htktAlloc = 0;
      let landAlloc = 0;
      let reserveAlloc = 0;
      let interestAlloc = 0;

      if (isSale) {
        const ratio = saleProgress / saleDenom;
        cpxdAlloc = meta.poolCPXD * ratio;
        gpmbAlloc = meta.poolGPMB * ratio;
        htktAlloc = meta.poolHTKT * ratio;
        landAlloc = meta.poolLand * ratio;
        reserveAlloc = meta.poolDuPhong * ratio;
        interestAlloc = meta.poolLaiVay * ratio;
      } else if (isRent && rentRate > 0) {
        cpxdAlloc = meta.poolCPXD / rentDuration;
        gpmbAlloc = meta.poolGPMB / rentDuration;
        htktAlloc = meta.poolHTKT / rentDuration;
        landAlloc = meta.poolLand / rentDuration;
        reserveAlloc = meta.poolDuPhong / rentDuration;
        interestAlloc = meta.poolLaiVay / rentDuration;
      }

      const sellingCost = isSale ? saleRevenue * saleCostRate : 0;
      const operatingCost = isRent ? rentRevenue * p.opCostRate : 0;

      const totalTaxCost =
        cpxdAlloc +
        sellingCost +
        operatingCost +
        gpmbAlloc +
        htktAlloc +
        landAlloc +
        reserveAlloc +
        interestAlloc;

      const taxableProfit = Math.max(0, revenue - totalTaxCost);
      const cit = taxableProfit * p.citRate;

      rows.push([
        t,                                      // A
        monthDate,                              // B
        year,                                   // C
        quarter,                                // D
        p.name,                                 // E
        p.method,                               // F
        p.area,                                 // G
        p.salePrice,                            // H
        p.rentPrice,                            // I
        p.cpxdUnit,                             // J
        p.occupancy,                            // K
        p.opCostRate,                           // L
        p.landArea,                             // M
        saleProgress || rentRate || 0,          // N
        saleRevenue,                            // O
        rentRevenue,                            // P
        revenue,                                // Q
        p.vatOut,                               // R
        vatOut,                                 // S
        cashFromCustomer,                       // T
        p.citRate,                              // U
        cpxdAlloc,                              // V
        sellingCost,                            // W
        operatingCost,                          // X
        gpmbAlloc,                              // Y
        htktAlloc,                              // Z
        landAlloc,                              // AA
        reserveAlloc,                           // AB
        interestAlloc,                          // AC
        totalTaxCost,                           // AD
        taxableProfit,                          // AE
        cit                                    // AF
      ]);
    });
  }

  if (rows.length) {
    sh.getRange(2, 1, rows.length, 32).setValues(rows);
  }

  FS02_formatSheet_(sh, endRow);
  FS02_hideZeroRevenueRows_(sh, endRow);
}

/***********************
 * DATA READERS
 ***********************/

function FS02V2_readInfo_(sheet) {
  const blockRow = FS02V2_findRowExact_(sheet, 'THONG_TIN_CHUNG');
  if (!blockRow) throw new Error('Không tìm thấy block THONG_TIN_CHUNG trong sheet 01. Kỹ thuật.');

  const info = {};
  for (let r = blockRow + 1; r <= sheet.getLastRow(); r++) {
    const label = String(sheet.getRange(r, 1).getDisplayValue() || '').trim();
    if (!label) break;
    if (FS02V2_isBlockMarker_(label)) break;
    info[label] = sheet.getRange(r, 2).getValue();
  }
  return info;
}

function FS02V2_readProducts_(sheet) {
  const block = FS02V2_getBlock_(sheet, 'SAN_PHAM');
  const out = [];

  block.rows.forEach(r => {
    const name = String(r[0] || '').trim();
    if (!name) return;

    out.push({
      name,
      method: String(r[1] || '').trim(),
      area: FS02V2_num_(r[2]),
      salePrice: FS02V2_num_(r[3]),
      rentPrice: FS02V2_num_(r[4]),
      cpxdUnit: FS02V2_num_(r[5]),
      vatOut: FS02V2_rate_(r[6]),
      citRate: FS02V2_rate_(r[7]),
      occupancy: FS02V2_rate_(r[8]),
      opCostRate: FS02V2_rate_(r[9]),
      note: String(r[10] || ''),
      landArea: FS02V2_num_(r[11])
    });
  });

  return out;
}

function FS02V2_readSchedules_(sheet) {
  const block = FS02V2_getBlock_(sheet, 'KE_HOACH_BAN_THU_TIEN');
  const out = [];

  block.rows.forEach(r => {
    const group = String(r[0] || '').trim();
    const product = String(r[1] || '').trim();
    if (!group || !product) return;

    out.push({
      group,
      product,
      dotCount: FS02V2_num_(r[2]),
      dot: FS02V2_num_(r[3]),
      start: FS02V2_num_(r[4]),
      duration: Math.max(1, FS02V2_num_(r[5])),
      rate: FS02V2_rate_(r[6]),
      note: String(r[7] || '')
    });
  });

  return out;
}

function FS02V2_readCommonCosts_(sheet) {
  const block = FS02V2_getBlock_(sheet, 'CHI_PHI_CHUNG');
  const out = {};

  block.rows.forEach(r => {
    const name = String(r[0] || '').trim();
    if (!name) return;
    out[FS02V2_key_(name)] = {
      name,
      beforeVat: FS02V2_num_(r[1]),
      vat: FS02V2_rate_(r[2]),
      afterVat: FS02V2_num_(r[3]),
      note: String(r[4] || ''),
      rate: FS02V2_rate_(r[5])
    };
  });

  return out;
}

function FS02V2_readTotals03_(sh03, endRow03) {
  const n = Math.max(0, endRow03 - 1);
  if (n <= 0) return { cpxd: 0, gpmb: 0, htkt: 0, duPhong: 0, laiVay: 0 };

  const values = sh03.getRange(2, 1, n, Math.min(35, sh03.getLastColumn())).getValues();

  return values.reduce((o, r) => {
    o.cpxd += FS02V2_num_(r[7]);      // H
    o.gpmb += FS02V2_num_(r[8]);      // I
    o.htkt += FS02V2_num_(r[10]);     // K
    o.duPhong += FS02V2_num_(r[12]);  // M
    o.laiVay += FS02V2_num_(r[25]);   // Z
    return o;
  }, { cpxd: 0, gpmb: 0, htkt: 0, duPhong: 0, laiVay: 0 });
}

/***********************
 * PRODUCT COST POOLS
 ***********************/

function FS02V2_buildProductMeta_(products, schedules, commonCosts, totals) {
  const out = {};

  const totalCPXDBase = FS02V2_sum_(products.map(p => p.area * p.cpxdUnit));
  const totalLandBase = FS02V2_sum_(products.map(p => p.landArea || p.area));
  const totalHTKTBase = totalCPXDBase || totalLandBase || 1;

  products.forEach(p => {
    const key = FS02V2_key_(p.name);
    const isSale = FS02V2_isSale_(p.method);
    const isRent = FS02V2_isRent_(p.method);

    const cpxdBase = p.area * p.cpxdUnit;
    const landBase = p.landArea || p.area || 0;

    const poolCPXD = totalCPXDBase ? totals.totalCPXD * cpxdBase / totalCPXDBase : 0;
    const poolGPMB = totalLandBase ? totals.totalGPMB * landBase / totalLandBase : 0;
    const poolHTKT = totalHTKTBase ? totals.totalHTKT * cpxdBase / totalHTKTBase : 0;
    const poolDuPhong = totalHTKTBase ? totals.totalDuPhong * cpxdBase / totalHTKTBase : 0;
    const poolLaiVay = totalHTKTBase ? totals.totalLaiVay * cpxdBase / totalHTKTBase : 0;

    const poolLand = FS02V2_getLandPoolForProduct_(p, commonCosts);

    const saleProgressTotal = FS02V2_totalSaleProgress_(p.name, schedules);
    const rentInfo = FS02V2_rentInfo_(p.name, schedules);

    out[key] = {
      isSale,
      isRent,
      poolCPXD,
      poolGPMB,
      poolHTKT,
      poolLand,
      poolDuPhong,
      poolLaiVay,
      totalSaleProgress: saleProgressTotal || 1,
      rentStart: rentInfo.start,
      rentDuration: rentInfo.duration || 1
    };
  });

  return out;
}

function FS02V2_getLandPoolForProduct_(p, commonCosts) {
  const nameKey = FS02V2_key_(p.name);

  if (FS02V2_hasAny_(nameKey, ['lienke', 'lien ke', 'liền kề', 'liên kề'])) {
    return FS02V2_getCostAmount_(commonCosts, [
      'Tiền SDĐ liền kề',
      'Tiền SDĐ đất liền kề',
      'Tiền SD đất liền kề',
      'Tiền sử dụng đất liền kề'
    ]);
  }

  if (FS02V2_hasAny_(nameKey, ['chungcu', 'chung cư', 'canho', 'căn hộ'])) {
    return FS02V2_getCostAmount_(commonCosts, [
      'Tiền thuê đất chung cư',
      'Tiền thuê đất Chung cư',
      'Tiền SDĐ chung cư',
      'Tiền sử dụng đất chung cư'
    ]);
  }

  if (FS02V2_hasAny_(nameKey, ['cho', 'chợ'])) {
    return FS02V2_getCostAmount_(commonCosts, [
      'Tiền thuê đất Chợ',
      'Tiền thuê đất chợ'
    ]);
  }

  if (FS02V2_hasAny_(nameKey, ['tmdv', 'thuongmai', 'thuong mai', 'thương mại'])) {
    return FS02V2_getCostAmount_(commonCosts, [
      'Tiền thuê đất TMDV',
      'Tiền thuê đất thương mại dịch vụ',
      'Tiền thuê đất TM-DV'
    ]);
  }

  return 0;
}

/***********************
 * SCHEDULE LOGIC
 ***********************/

function FS02V2_getSaleProgress_(productName, monthNo, schedules) {
  return schedules
    .filter(s =>
      FS02V2_isGroup_(s.group, ['Thu tiền', 'Thu tien']) &&
      FS02V2_key_(s.product) === FS02V2_key_(productName) &&
      monthNo >= s.start &&
      monthNo < s.start + s.duration
    )
    .reduce((sum, s) => sum + (s.duration > 0 ? s.rate / s.duration : 0), 0);
}

function FS02V2_getRentActiveRate_(productName, monthNo, schedules) {
  return schedules
    .filter(s =>
      FS02V2_isGroup_(s.group, ['Thu tiền', 'Thu tien']) &&
      FS02V2_key_(s.product) === FS02V2_key_(productName) &&
      monthNo >= s.start &&
      monthNo < s.start + s.duration
    )
    .reduce((sum, s) => sum + s.rate, 0);
}

function FS02V2_totalSaleProgress_(productName, schedules) {
  return schedules
    .filter(s =>
      FS02V2_isGroup_(s.group, ['Thu tiền', 'Thu tien']) &&
      FS02V2_key_(s.product) === FS02V2_key_(productName)
    )
    .reduce((sum, s) => sum + s.rate, 0);
}

function FS02V2_rentInfo_(productName, schedules) {
  const rows = schedules
    .filter(s =>
      FS02V2_isGroup_(s.group, ['Thu tiền', 'Thu tien']) &&
      FS02V2_key_(s.product) === FS02V2_key_(productName)
    )
    .sort((a, b) => a.start - b.start);

  if (!rows.length) return { start: 1, duration: 1 };

  return {
    start: rows[0].start,
    duration: rows.reduce((maxEnd, s) => Math.max(maxEnd, s.start + s.duration - rows[0].start), 0)
  };
}

/***********************
 * GENERIC HELPERS
 ***********************/

function FS02V2_getBlock_(sheet, blockName) {
  const blockRow = FS02V2_findRowExact_(sheet, blockName);
  if (!blockRow) throw new Error('Không tìm thấy block ' + blockName);

  const headerRow = blockRow + 1;
  const lastCol = FS02V2_lastColInRow_(sheet, headerRow);
  const rows = [];
  let blank = 0;

  for (let r = blockRow + 2; r <= sheet.getLastRow(); r++) {
    const first = String(sheet.getRange(r, 1).getDisplayValue() || '').trim();
    if (FS02V2_isBlockMarker_(first)) break;

    const row = sheet.getRange(r, 1, 1, lastCol).getValues()[0];
    const hasData = row.some(v => String(v || '').trim() !== '');

    if (!hasData) {
      blank++;
      if (blank >= 3) break;
      continue;
    }

    blank = 0;
    rows.push(row);
  }

  return { blockRow, headerRow, rows };
}

function FS02V2_findRowExact_(sheet, text) {
  const values = sheet.getDataRange().getDisplayValues();
  const target = FS02V2_key_(text);

  for (let r = 0; r < values.length; r++) {
    if (values[r].some(v => FS02V2_key_(v) === target)) return r + 1;
  }

  return null;
}

function FS02V2_lastColInRow_(sheet, rowNo) {
  const values = sheet.getRange(rowNo, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  let last = 1;
  values.forEach((v, i) => {
    if (String(v || '').trim() !== '') last = i + 1;
  });
  return last;
}

function FS02V2_prepareSheet_(sh, endRow, lastCol) {
  sh.showRows(1, sh.getMaxRows());
  sh.showColumns(1, sh.getMaxColumns());

  if (sh.getMaxRows() < endRow) {
    sh.insertRowsAfter(sh.getMaxRows(), endRow - sh.getMaxRows());
  }

  if (sh.getMaxColumns() < lastCol) {
    sh.insertColumnsAfter(sh.getMaxColumns(), lastCol - sh.getMaxColumns());
  }

  sh.clearContents();
  sh.clearFormats();
}

function FS02V2_getSheetByPrefix_(ss, prefix) {
  return ss.getSheets().find(s => String(s.getName()).trim().startsWith(prefix));
}

function FS02V2_isBlockMarker_(v) {
  return ['THONG_TIN_CHUNG', 'CHI_PHI_CHUNG', 'SAN_PHAM', 'KE_HOACH_BAN_THU_TIEN', 'TIEN_DO_CHI_PHI']
    .includes(String(v || '').trim());
}

function FS02V2_isGroup_(value, names) {
  const k = FS02V2_key_(value);
  return names.some(n => k === FS02V2_key_(n));
}

function FS02V2_isSale_(method) {
  return FS02V2_key_(method) === FS02V2_key_('Bán');
}

function FS02V2_isRent_(method) {
  return FS02V2_key_(method) === FS02V2_key_('Cho thuê');
}

function FS02V2_getCostAmount_(costMap, names) {
  for (const name of names) {
    const item = costMap[FS02V2_key_(name)];
    if (item) return item.beforeVat;
  }
  return 0;
}

function FS02V2_getCostRate_(costMap, names) {
  for (const name of names) {
    const item = costMap[FS02V2_key_(name)];
    if (item) return item.rate;
  }
  return 0;
}

function FS02V2_hasAny_(key, terms) {
  const k = FS02V2_key_(key);
  return terms.some(t => k.indexOf(FS02V2_key_(t)) >= 0);
}

function FS02V2_addMonths_(dateValue, months) {
  const d = new Date(dateValue);
  const out = new Date(d.getFullYear(), d.getMonth() + months, d.getDate());
  return out;
}

function FS02V2_sum_(arr) {
  return arr.reduce((s, v) => s + (Number(v) || 0), 0);
}

function FS02V2_num_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  const s = String(v || '').replace(/\s/g, '').replace(/,/g, '').replace(/%/g, '');
  const n = Number(s);
  return isFinite(n) ? n : 0;
}

function FS02V2_rate_(v) {
  if (v === '' || v === null || typeof v === 'undefined') return 0;
  if (typeof v === 'number') return v > 1 ? v / 100 : v;

  const s = String(v || '').trim();
  if (!s) return 0;

  const n = Number(s.replace('%', '').replace(',', '.'));
  if (!isFinite(n)) return 0;

  return n > 1 ? n / 100 : n;
}

function FS02V2_key_(v) {
  return String(v || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/²/g, '2')
    .replace(/\^2/g, '2')
    .replace(/m\s*2/g, 'm2')
    .replace(/m\s*²/g, 'm2')
    .replace(/[^a-z0-9]/g, '');
}

/***********************
 * FORMAT / HIDE
 * Giữ lại tên hàm cũ để tương thích các file khác.
 ***********************/

function FS02_hideZeroRevenueRows_(sh, endRow) {
  if (endRow <= 2) return;

  sh.showRows(1, sh.getMaxRows());

  const values = sh.getRange(2, 17, endRow - 1, 1).getValues();
  let startHide = null;
  let countHide = 0;

  for (let i = 0; i < values.length; i++) {
    const rowNo = i + 2;
    const revenue = Number(values[i][0]) || 0;

    if (revenue === 0) {
      if (startHide === null) {
        startHide = rowNo;
        countHide = 1;
      } else {
        countHide++;
      }
    } else if (startHide !== null) {
      sh.hideRows(startHide, countHide);
      startHide = null;
      countHide = 0;
    }
  }

  if (startHide !== null) sh.hideRows(startHide, countHide);
}

function FS02_formatSheet_(sh, endRow) {
  const lastCol = 32;

  sh.setFrozenRows(1);

  sh.getRange(1, 1, 1, lastCol)
    .setFontWeight('bold')
    .setBackground('#d9ead3')
    .setHorizontalAlignment('center')
    .setWrap(true);

  sh.getRange(1, 1, endRow, lastCol)
    .setFontFamily('Arial')
    .setFontSize(10)
    .setVerticalAlignment('middle');

  if (endRow > 1) {
    const rows = endRow - 1;

    sh.getRange(2, 1, rows, 1).setNumberFormat('0');
    sh.getRange(2, 2, rows, 1).setNumberFormat('dd/mm/yyyy');
    sh.getRange(2, 3, rows, 1).setNumberFormat('0');

    sh.getRange(2, 7, rows, 1).setNumberFormat('#,##0.00');
    sh.getRange(2, 8, rows, 3).setNumberFormat('#,##0');
    sh.getRange(2, 11, rows, 2).setNumberFormat('0.00%');
    sh.getRange(2, 13, rows, 1).setNumberFormat('#,##0.00');
    sh.getRange(2, 14, rows, 1).setNumberFormat('0.00%');

    sh.getRange(2, 15, rows, 3).setNumberFormat('#,##0');
    sh.getRange(2, 18, rows, 1).setNumberFormat('0.00%');
    sh.getRange(2, 19, rows, 2).setNumberFormat('#,##0');
    sh.getRange(2, 21, rows, 1).setNumberFormat('0.00%');
    sh.getRange(2, 22, rows, 11).setNumberFormat('#,##0');
  }

  sh.autoResizeColumns(1, lastCol);
}
