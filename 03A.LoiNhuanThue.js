function FS_lapSheet03A() { return FS_lapSheet03A_voiLaiVay_(null); }

function FS_lapSheet03A_voiLaiVay_(interest) {
  const ss = SpreadsheetApp.getActive();
  const rev = ss.getSheetByName(FS_CFG.SHEETS.REVENUE);
  const cost = ss.getSheetByName(FS_CFG.SHEETS.COST);
  if (!rev || !cost) throw new Error('Cần lập 02. Doanh thu và 03. Chi phí & Vốn trước.');

  const revRows = rev.getLastRow() > 1 ? rev.getRange(2, 1, rev.getLastRow() - 1, 20).getValues() : [];
  const costRows = cost.getLastRow() > 1 ? cost.getRange(2, 1, cost.getLastRow() - 1, 20).getValues() : [];
  const costByKey = {};
  costRows.forEach(r => costByKey[FS_factKey_(r[0], r[4])] = r);

  const totals = {};
  revRows.forEach(r => {
    const code = String(r[4]);
    if (!totals[code]) totals[code] = { saleRevenue: 0, rentActiveMonths: 0, capitalPool: 0 };
    totals[code].saleRevenue += FS_num_(r[13]);
    if (FS_num_(r[14]) > 0) totals[code].rentActiveMonths++;
  });
  costRows.forEach(r => {
    const code = String(r[4]);
    if (!totals[code]) totals[code] = { saleRevenue: 0, rentActiveMonths: 0, capitalPool: 0 };
    totals[code].capitalPool += FS_num_(r[9]) + FS_num_(r[10]) + FS_num_(r[11]) + FS_num_(r[12]) + FS_num_(r[16]);
  });

  const investmentByMonth = {};
  costRows.forEach(r => {
    const t = FS_num_(r[0]);
    investmentByMonth[t] = (investmentByMonth[t] || 0) + FS_num_(r[9]) + FS_num_(r[10]) + FS_num_(r[11]) + FS_num_(r[12]) + FS_num_(r[16]);
  });

  const rows = [];
  revRows.forEach(r => {
    const t = FS_num_(r[0]);
    const code = String(r[4]);
    const c = costByKey[FS_factKey_(t, code)] || [];
    const revenue = FS_num_(r[15]);
    const saleRevenue = FS_num_(r[13]);
    const rentRevenue = FS_num_(r[14]);
    const capitalPool = totals[code].capitalPool;

    const capitalRecognized = r[6] === 'Bán'
      ? capitalPool * FS_ratio_(saleRevenue, totals[code].saleRevenue)
      : (rentRevenue > 0 ? capitalPool / Math.max(1, totals[code].rentActiveMonths) : 0);

    const currentInvest = FS_num_(c[9]) + FS_num_(c[10]) + FS_num_(c[11]) + FS_num_(c[12]) + FS_num_(c[16]);
    const interestMonth = interest ? FS_num_(interest[t - 1]) : 0;
    const interestAlloc = interestMonth * FS_ratio_(currentInvest, investmentByMonth[t]);

    const selling = FS_num_(c[13]);
    const operating = FS_num_(c[14]);
    const maintenance = FS_num_(c[15]);
    const taxCost = capitalRecognized + selling + operating + maintenance;
    const pbt = revenue - taxCost - interestAlloc;
    const taxable = Math.max(0, pbt);
    const citRate = FS_rate_(r[19]);
    const cit = taxable * citRate;
    const pat = pbt - cit;

    rows.push([
      r[0], r[1], r[2], r[3], code, r[5], r[6], revenue,
      capitalRecognized, selling, operating, maintenance, taxCost,
      interestAlloc, pbt, taxable, citRate, cit, pat
    ]);
  });

  const sh = FS_getOrCreateSheet_(ss, FS_CFG.SHEETS.PROFIT);
  FS_resetSheet_(sh, rows.length + 1, 19);
  sh.getRange(1, 1, 1, 19).setValues([[
    'Tháng số', 'Tháng', 'Năm', 'Quý', 'Mã SP', 'Tên sản phẩm', 'Loại hình',
    'Doanh thu trước VAT', 'Giá vốn đầu tư ghi nhận', 'Chi phí bán hàng',
    'Chi phí vận hành', 'Chi phí bảo trì', 'Tổng chi phí tính thuế trước lãi vay',
    'Lãi vay phân bổ', 'Lợi nhuận trước thuế', 'Thu nhập chịu thuế',
    'Thuế suất TNDN (%)', 'Thuế TNDN', 'LNST'
  ]]);
  if (rows.length) sh.getRange(2, 1, rows.length, 19).setValues(rows);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(7);
  sh.getRange(1, 1, 1, 19).setFontWeight('bold').setBackground('#e2f0d9').setWrap(true);
  if (rows.length) {
    sh.getRange(2, 2, rows.length, 1).setNumberFormat('MM/yyyy');
    sh.getRange(2, 8, rows.length, 12).setNumberFormat('#,##0.00');
    sh.getRange(2, 17, rows.length, 1).setNumberFormat('0.00%');
  }
  sh.autoResizeColumns(1, 19);
  return rows;
}
