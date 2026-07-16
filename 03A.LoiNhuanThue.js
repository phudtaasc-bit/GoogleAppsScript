function FS_lapSheet03A() { return FS_lapSheet03A_voiLaiVay_(null); }

function FS_lapSheet03A_voiLaiVay_(interest) {
  const ss = SpreadsheetApp.getActive();
  const tech = FS_getSheet_(ss, FS_CFG.SHEETS.TECH, FS_CFG.SHEETS.TECH_LEGACY);
  const rev = ss.getSheetByName(FS_CFG.SHEETS.REVENUE);
  const cost = ss.getSheetByName(FS_CFG.SHEETS.COST);
  if (!tech || !rev || !cost) throw new Error('Cần lập 01A. Kỹ thuật, 02. Doanh thu và 03. Chi phí & Vốn trước.');

  const revRows = rev.getLastRow() > 1 ? rev.getRange(2, 1, rev.getLastRow() - 1, 20).getValues() : [];
  const costRows = cost.getLastRow() > 1 ? cost.getRange(2, 1, cost.getLastRow() - 1, 21).getValues() : [];
  const productRows = FS_readBlock_(tech, 'SAN_PHAM');
  const leaseYearsByCode = {};
  productRows.forEach(r => {
    const name = String(r[0] || '').trim();
    if (name) leaseYearsByCode[FS_maSanPhamGoc_(name)] = Math.max(0, FS_num_(r[11]));
  });

  const costByKey = {};
  costRows.forEach(r => costByKey[FS_factKey_(r[0], r[4])] = r);

  const totals = {};
  revRows.forEach(r => {
    const code = String(r[4]);
    if (!totals[code]) totals[code] = {
      type: r[6], saleRevenue: 0, firstRentMonth: 0,
      baseCapitalPool: 0, landUsePool: 0, landRentPool: 0
    };
    totals[code].saleRevenue += FS_num_(r[13]);
    if (!totals[code].firstRentMonth && FS_num_(r[14]) > 0) totals[code].firstRentMonth = FS_num_(r[0]);
  });

  costRows.forEach(r => {
    const code = String(r[4]);
    if (!totals[code]) totals[code] = {
      type: r[6], saleRevenue: 0, firstRentMonth: 0,
      baseCapitalPool: 0, landUsePool: 0, landRentPool: 0
    };
    totals[code].baseCapitalPool += FS_num_(r[9]) + FS_num_(r[10]) + FS_num_(r[13]) + FS_num_(r[17]);
    totals[code].landUsePool += FS_num_(r[11]);
    totals[code].landRentPool += FS_num_(r[12]);
  });

  const investmentByMonth = {};
  costRows.forEach(r => {
    const t = FS_num_(r[0]);
    investmentByMonth[t] = (investmentByMonth[t] || 0) +
      FS_num_(r[9]) + FS_num_(r[10]) + FS_num_(r[11]) + FS_num_(r[12]) + FS_num_(r[13]) + FS_num_(r[17]);
  });

  const rows = [];
  revRows.forEach(r => {
    const t = FS_num_(r[0]);
    const code = String(r[4]);
    const c = costByKey[FS_factKey_(t, code)] || [];
    const meta = totals[code];
    const revenue = FS_num_(r[15]);
    const saleRevenue = FS_num_(r[13]);
    const rentRevenue = FS_num_(r[14]);

    let baseCapitalRecognized = 0;
    let landUseRecognized = 0;
    let landRentRecognized = 0;

    if (r[6] === 'Bán') {
      const saleRatio = FS_ratio_(saleRevenue, meta.saleRevenue);
      baseCapitalRecognized = meta.baseCapitalPool * saleRatio;
      landUseRecognized = meta.landUsePool * saleRatio;
      landRentRecognized = meta.landRentPool * saleRatio;
    } else {
      const active = rentRevenue > 0;
      const activeMonths = Math.max(1, FS_num_(leaseYearsByCode[code]) * 12);
      if (active) {
        const rentRows = revRows.filter(x => String(x[4]) === code && FS_num_(x[14]) > 0).length || 1;
        baseCapitalRecognized = meta.baseCapitalPool / rentRows;
        landRentRecognized = meta.landRentPool / activeMonths;
      }
    }

    const currentInvest = FS_num_(c[9]) + FS_num_(c[10]) + FS_num_(c[11]) + FS_num_(c[12]) + FS_num_(c[13]) + FS_num_(c[17]);
    const interestMonth = interest ? FS_num_(interest[t - 1]) : 0;
    const interestAlloc = interestMonth * FS_ratio_(currentInvest, investmentByMonth[t]);

    const selling = FS_num_(c[14]);
    const operating = FS_num_(c[15]);
    const maintenance = FS_num_(c[16]);
    const capitalRecognized = baseCapitalRecognized + landUseRecognized + landRentRecognized;
    const taxCost = capitalRecognized + selling + operating + maintenance;
    const pbt = revenue - taxCost - interestAlloc;
    const taxable = Math.max(0, pbt);
    const citRate = FS_rate_(r[19]);
    const cit = taxable * citRate;
    const pat = pbt - cit;

    rows.push([
      r[0], r[1], r[2], r[3], code, r[5], r[6], revenue,
      baseCapitalRecognized, landUseRecognized, landRentRecognized,
      selling, operating, maintenance, taxCost,
      interestAlloc, pbt, taxable, citRate, cit, pat
    ]);
  });

  const sh = FS_getOrCreateSheet_(ss, FS_CFG.SHEETS.PROFIT);
  FS_resetSheet_(sh, rows.length + 1, 21);
  sh.getRange(1, 1, 1, 21).setValues([[
    'Tháng số', 'Tháng', 'Năm', 'Quý', 'Mã SP', 'Tên sản phẩm', 'Loại hình',
    'Doanh thu trước VAT', 'Giá vốn XD/GPMB/HTKT/Dự phòng',
    'Tiền SDĐ hạch toán', 'Tiền thuê đất hạch toán/phân bổ',
    'Chi phí bán hàng', 'Chi phí vận hành', 'Chi phí bảo trì',
    'Tổng chi phí tính thuế trước lãi vay', 'Lãi vay phân bổ',
    'Lợi nhuận trước thuế', 'Thu nhập chịu thuế',
    'Thuế suất TNDN (%)', 'Thuế TNDN', 'LNST'
  ]]);
  if (rows.length) sh.getRange(2, 1, rows.length, 21).setValues(rows);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(7);
  sh.getRange(1, 1, 1, 21).setFontWeight('bold').setBackground('#e2f0d9').setWrap(true);
  if (rows.length) {
    sh.getRange(2, 2, rows.length, 1).setNumberFormat('MM/yyyy');
    sh.getRange(2, 8, rows.length, 11).setNumberFormat('#,##0');
    sh.getRange(2, 19, rows.length, 1).setNumberFormat('0.00%');
    sh.getRange(2, 20, rows.length, 2).setNumberFormat('#,##0');
  }
  sh.autoResizeColumns(1, 21);
  return rows;
}
