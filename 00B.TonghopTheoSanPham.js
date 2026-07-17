
/**
 * 00B.TonghopTheoSanPham.js
 *
 * Phân tích hiệu quả đầu tư riêng từng sản phẩm trên mảng dữ liệu.
 * Không ghi đè và không chạy lại các sheet lõi 02, 03, 03A, 04, 04A.
 *
 * Logic:
 * - Khi phân tích một sản phẩm, chỉ giữ doanh thu và chi phí của sản phẩm đó.
 * - VAT phải nộp được tính lại riêng theo chuỗi VAT đầu ra/đầu vào của sản phẩm.
 * - Thuế TNDN và lãi vay phân bổ lấy từ 03A; không hội tụ lại lãi vay.
 * - Giải ngân vay/trả gốc gốc của dự án được phân bổ theo tỷ trọng chi phí sản phẩm,
 *   nhằm xác định FCFE tham chiếu mà không làm thay đổi mô hình gốc.
 */

const FS00B_CFG = {
  summarySheet: '00. Tổng hợp',
  revenueSheet: '02. Doanh thu',
  costSheet: '03. Chi phí & Vốn',
  taxSheet: '03A. Lợi nhuận & Thuế',
  cashSheet: '04. Dòng tiền & Tài trợ',
  firstProductColumn: 5, // E
  billion: 1e9,
  tolerance: 1
};

function FS00B_phanTichHieuQuaTheoSanPham() {
  const ss = SpreadsheetApp.getActive();
  const sheets = FS00B_getRequiredSheets_(ss);

  const revenue = FS00B_readTable_(sheets.revenue);
  const cost = FS00B_readTable_(sheets.cost);
  const tax = FS00B_readTable_(sheets.tax);
  const cash = FS00B_readTable_(sheets.cash);

  FS00B_require_(revenue, [
    'thangso', 'thang', 'masp', 'tensanpham', 'nhom',
    'tongdoanhthutruocvat', 'vatdaura', 'dongtienkhachhang'
  ], sheets.revenue.getName());

  FS00B_require_(cost, [
    'thangso', 'masp', 'tongchitruocvat', 'vatdauvao', 'tongchisauvat',
    'xdtbtruocvat', 'gpmbtruocvat', 'tiensddtruocvat',
    'tienthuedattruocvat', 'htkttruocvat',
    'chiphibanhangtruocvat', 'chiphivanhanhtruocvat',
    'chiphibaotritruocvat', 'chiphiduphongtruocvat'
  ], sheets.cost.getName());

  FS00B_require_(tax, [
    'thangso', 'masp', 'laivayphanbo', 'thuetndn', 'lnst'
  ], sheets.tax.getName());

  FS00B_require_(cash, [
    'thangso', 'thang', 'giainganvay', 'tragoc'
  ], sheets.cash.getName());

  const products = FS00B_getProducts_(revenue);
  if (!products.length) throw new Error('Không tìm thấy sản phẩm trong Sheet 02.');

  const layout = FS00C_chuanBiCotSanPham_(sheets.summary, products);
  const projectRates = FS00B_readDiscountRates_(sheets.summary);

  const projectFinance = FS00B_buildProjectFinance_(cash);
  const allProductCosts = FS00B_buildAllProductCostMap_(cost, products);
  const totalCostByMonth = FS00B_sumCostByMonth_(allProductCosts);
  const globalCostShares = FS00B_globalCostShares_(allProductCosts, products);

  const results = products.map(product => FS00B_analyzeProduct_({
    product,
    revenue,
    cost,
    tax,
    projectFinance,
    totalCostByMonth,
    globalCostShare: globalCostShares[product.code] || 0,
    projectRates
  }));

  FS00B_writeResults_(sheets.summary, layout, products, results);
  SpreadsheetApp.flush();

  ss.toast(
    'Đã cập nhật NPV, IRR và các chỉ tiêu hiệu quả cho ' + products.length + ' sản phẩm.',
    'Phân tích theo sản phẩm',
    6
  );

  return results;
}

function FS00B_analyzeProduct_(ctx) {
  const code = ctx.product.code;
  const revRows = FS00B_rowsByCode_(ctx.revenue, code);
  const costRows = FS00B_rowsByCode_(ctx.cost, code);
  const taxRows = FS00B_rowsByCode_(ctx.tax, code);

  const monthCount = ctx.projectFinance.months.length;
  const revMap = FS00B_indexRowsByMonth_(revRows, ctx.revenue, 'thangso');
  const costMap = FS00B_indexRowsByMonth_(costRows, ctx.cost, 'thangso');
  const taxMap = FS00B_indexRowsByMonth_(taxRows, ctx.tax, 'thangso');

  const dates = [];
  const fcff = [];
  const fcfe = [];

  let vatCredit = 0;
  let cumulativeDebt = 0;
  let peakDebt = 0;

  const sums = {
    revenueWithVat: 0,
    totalCostWithVat: 0,
    coreInvestmentWithVat: 0,
    sellingWithVat: 0,
    operatingWithVat: 0,
    maintenanceWithVat: 0,
    lnst: 0,
    interest: 0,
    cit: 0,
    vatPayable: 0
  };

  for (let i = 0; i < monthCount; i++) {
    const monthNo = ctx.projectFinance.months[i];
    const date = ctx.projectFinance.dates[i];
    const r = revMap[monthNo] || null;
    const c = costMap[monthNo] || null;
    const t = taxMap[monthNo] || null;

    const customerCash = r ? FS00B_num_(r[ctx.revenue.index.dongtienkhachhang]) : 0;
    const vatOut = r ? FS00B_num_(r[ctx.revenue.index.vatdaura]) : 0;

    const costBeforeVat = c ? FS00B_num_(c[ctx.cost.index.tongchitruocvat]) : 0;
    const vatIn = c ? FS00B_num_(c[ctx.cost.index.vatdauvao]) : 0;
    const costAfterVat = c ? FS00B_num_(c[ctx.cost.index.tongchisauvat]) : 0;

    const cit = t ? FS00B_num_(t[ctx.tax.index.thuetndn]) : 0;
    const lnst = t ? FS00B_num_(t[ctx.tax.index.lnst]) : 0;
    const interest = t ? FS00B_num_(t[ctx.tax.index.laivayphanbo]) : 0;

    const availableVatCredit = vatCredit + vatIn;
    const vatPayable = Math.max(0, vatOut - availableVatCredit);
    vatCredit = Math.max(0, availableVatCredit - vatOut);

    const productFcff = customerCash - costAfterVat - vatPayable - cit;

    const monthTotalCost = FS00B_num_(ctx.totalCostByMonth[monthNo]);
    const monthShare = monthTotalCost > FS00B_CFG.tolerance
      ? costAfterVat / monthTotalCost
      : ctx.globalCostShare;

    const allocatedDraw = ctx.projectFinance.draw[i] * monthShare;
    const allocatedPrincipal = ctx.projectFinance.principal[i] * monthShare;
    const productFcfe = productFcff + allocatedDraw - allocatedPrincipal;

    cumulativeDebt += allocatedDraw - allocatedPrincipal;
    cumulativeDebt = Math.max(0, cumulativeDebt);
    peakDebt = Math.max(peakDebt, cumulativeDebt);

    dates.push(date);
    fcff.push(productFcff);
    fcfe.push(productFcfe);

    sums.revenueWithVat += customerCash;
    sums.totalCostWithVat += costAfterVat;
    sums.lnst += lnst;
    sums.interest += interest;
    sums.cit += cit;
    sums.vatPayable += vatPayable;

    if (c) {
      const ratio = costBeforeVat > FS00B_CFG.tolerance
        ? costAfterVat / costBeforeVat
        : 1;

      const selling = FS00B_num_(c[ctx.cost.index.chiphibanhangtruocvat]) * ratio;
      const operating = FS00B_num_(c[ctx.cost.index.chiphivanhanhtruocvat]) * ratio;
      const maintenance = FS00B_num_(c[ctx.cost.index.chiphibaotritruocvat]) * ratio;

      const coreBeforeVat =
        FS00B_num_(c[ctx.cost.index.xdtbtruocvat]) +
        FS00B_num_(c[ctx.cost.index.gpmbtruocvat]) +
        FS00B_num_(c[ctx.cost.index.tiensddtruocvat]) +
        FS00B_num_(c[ctx.cost.index.tienthuedattruocvat]) +
        FS00B_num_(c[ctx.cost.index.htkttruocvat]) +
        FS00B_num_(c[ctx.cost.index.chiphiduphongtruocvat]);

      sums.sellingWithVat += selling;
      sums.operatingWithVat += operating;
      sums.maintenanceWithVat += maintenance;
      sums.coreInvestmentWithVat += coreBeforeVat * ratio;
    }
  }

  // Tổng vốn đầu tư sản phẩm gồm chi phí đầu tư cốt lõi sau VAT và lãi vay đã phân bổ cố định.
  sums.coreInvestmentWithVat += sums.interest;

  return {
    code: ctx.product.code,
    name: ctx.product.name,
    group: ctx.product.group,
    revenueWithVat: sums.revenueWithVat,
    totalCostWithVat: sums.totalCostWithVat + sums.interest,
    coreInvestmentWithVat: sums.coreInvestmentWithVat,
    sellingWithVat: sums.sellingWithVat,
    operatingWithVat: sums.operatingWithVat,
    maintenanceWithVat: sums.maintenanceWithVat,
    lnst: sums.lnst,
    npvProject: FS00B_xnpv_(ctx.projectRates.wacc, fcff, dates),
    irrProject: FS00B_xirr_(fcff, dates),
    paybackProject: FS00B_sustainablePayback_(fcff),
    npvEquity: FS00B_xnpv_(ctx.projectRates.costOfEquity, fcfe, dates),
    irrEquity: FS00B_xirr_(fcfe, dates),
    paybackEquity: FS00B_sustainablePayback_(fcfe),
    peakDebt: peakDebt,
    totalInterest: sums.interest,
    totalCit: sums.cit,
    totalVatPayable: sums.vatPayable
  };
}

function FS00B_writeResults_(summary, layout, products, results) {
  const rowMap = FS00B_summaryRows_(summary);
  const resultByCode = {};
  results.forEach(r => resultByCode[r.code] = r);

  const labels = {
    totalRevenue: ['Tổng doanh thu có VAT'],
    totalCost: ['Tổng chi phí có VAT'],
    coreInvestment: ['Tổng vốn đầu tư dự án'],
    selling: ['Chi phí bán hàng'],
    operating: ['Chi phí vận hành'],
    maintenance: ['Chi phí bảo trì'],
    lnst: ['Lợi nhuận sau thuế'],
    npvProject: ['NPV dự án'],
    irrProject: ['IRR dự án'],
    paybackProject: ['Thời gian hoàn vốn dự án'],
    npvEquity: ['NPV vốn CSH'],
    irrEquity: ['IRR vốn CSH'],
    paybackEquity: ['Thời gian hoàn vốn - Vốn CSH', 'Thời gian hoàn vốn vốn CSH'],
    peakDebt: ['Đỉnh dư nợ vay'],
    interest: ['Tổng lãi vay'],
    cit: ['Tổng Thuế TNDN', 'Tổng thuế TNDN'],
    vat: ['Tổng VAT phải nộp']
  };

  products.forEach((product, pIndex) => {
    const col = layout.firstProductColumn + pIndex;
    const r = resultByCode[product.code];
    if (!r) return;

    FS00B_setByAliases_(summary, rowMap, labels.totalRevenue, col, r.revenueWithVat / FS00B_CFG.billion);
    FS00B_setByAliases_(summary, rowMap, labels.totalCost, col, r.totalCostWithVat / FS00B_CFG.billion);
    FS00B_setByAliases_(summary, rowMap, labels.coreInvestment, col, r.coreInvestmentWithVat / FS00B_CFG.billion);
    FS00B_setByAliases_(summary, rowMap, labels.selling, col, r.sellingWithVat / FS00B_CFG.billion);
    FS00B_setByAliases_(summary, rowMap, labels.operating, col, r.operatingWithVat / FS00B_CFG.billion);
    FS00B_setByAliases_(summary, rowMap, labels.maintenance, col, r.maintenanceWithVat / FS00B_CFG.billion);
    FS00B_setByAliases_(summary, rowMap, labels.lnst, col, r.lnst / FS00B_CFG.billion);

    FS00B_setByAliases_(summary, rowMap, labels.npvProject, col, r.npvProject / FS00B_CFG.billion);
    FS00B_setByAliases_(summary, rowMap, labels.irrProject, col, r.irrProject);
    FS00B_setByAliases_(summary, rowMap, labels.paybackProject, col, r.paybackProject);

    FS00B_setByAliases_(summary, rowMap, labels.npvEquity, col, r.npvEquity / FS00B_CFG.billion);
    FS00B_setByAliases_(summary, rowMap, labels.irrEquity, col, r.irrEquity);
    FS00B_setByAliases_(summary, rowMap, labels.paybackEquity, col, r.paybackEquity);

    FS00B_setByAliases_(summary, rowMap, labels.peakDebt, col, r.peakDebt / FS00B_CFG.billion);
    FS00B_setByAliases_(summary, rowMap, labels.interest, col, r.totalInterest / FS00B_CFG.billion);
    FS00B_setByAliases_(summary, rowMap, labels.cit, col, r.totalCit / FS00B_CFG.billion);
    FS00B_setByAliases_(summary, rowMap, labels.vat, col, r.totalVatPayable / FS00B_CFG.billion);

    // Dòng doanh thu chi tiết: chỉ sản phẩm tương ứng có giá trị, sản phẩm khác bằng 0.
    products.forEach(detailProduct => {
      const aliases = [
        'Phần ' + detailProduct.name,
        'Phần ' + detailProduct.name + (detailProduct.group ? ' - ' + detailProduct.group : '')
      ];
      FS00B_setByAliases_(
        summary,
        rowMap,
        aliases,
        col,
        detailProduct.code === product.code ? r.revenueWithVat / FS00B_CFG.billion : 0
      );
    });
  });

  const lastProductCol = layout.firstProductColumn + products.length - 1;
  const lastRow = FS00B_lastSectionRow_(summary);
  if (lastRow >= layout.headerRow + 1) {
    summary.getRange(layout.headerRow + 1, layout.firstProductColumn, lastRow - layout.headerRow, products.length)
      .setNumberFormat('#,##0.0');

    FS00B_formatPercentRows_(summary, rowMap, labels.irrProject, layout.firstProductColumn, products.length);
    FS00B_formatPercentRows_(summary, rowMap, labels.irrEquity, layout.firstProductColumn, products.length);
    FS00B_formatPaybackRows_(summary, rowMap, labels.paybackProject, layout.firstProductColumn, products.length);
    FS00B_formatPaybackRows_(summary, rowMap, labels.paybackEquity, layout.firstProductColumn, products.length);
  }

  summary.autoResizeColumns(layout.firstProductColumn, products.length);
  summary.setColumnWidths(layout.firstProductColumn, products.length, 125);
  summary.setColumnWidth(layout.noteColumn, 210);
}

function FS00B_getRequiredSheets_(ss) {
  const names = FS00B_CFG;
  const result = {
    summary: ss.getSheetByName(names.summarySheet),
    revenue: ss.getSheetByName(names.revenueSheet),
    cost: ss.getSheetByName(names.costSheet),
    tax: ss.getSheetByName(names.taxSheet),
    cash: ss.getSheetByName(names.cashSheet)
  };
  Object.keys(result).forEach(key => {
    if (!result[key]) throw new Error('Không tìm thấy sheet: ' + names[key + 'Sheet']);
  });
  return result;
}

function FS00B_readTable_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error('Sheet không có dữ liệu: ' + sheet.getName());

  const index = {};
  values[0].forEach((header, col) => {
    index[FS00B_key_(header)] = col;
  });

  // Alias có khoảng trắng được chuẩn hóa lại cho dễ đọc ở các hàm trên.
  index.vatdaura = FS00B_findHeaderIndex_(index, ['VAT đầu ra']);
  index.dongtienkhachhang = FS00B_findHeaderIndex_(index, [
    'Dòng tiền khách hàng',
    'Dòng tiền thu khách hàng'
  ]);
  index.nhom = FS00B_findHeaderIndex_(index, [
    'Nhóm',
    'Loại hình'
  ]);
  index.tongdoanhthutruocvat = FS00B_findHeaderIndex_(index, [
    'Tổng doanh thu trước VAT'
  ]);
  index.tongchitruocvat = index[FS00B_key_('Tổng chi trước VAT')];
  index.vatdauvao = index[FS00B_key_('VAT đầu vào')];
  index.tongchisauvat = index[FS00B_key_('Tổng chi sau VAT')];
  index.xdtbtruocvat = index[FS00B_key_('XD/TB trước VAT')];
  index.gpmbtruocvat = index[FS00B_key_('GPMB trước VAT')];
  index.tiensddtruocvat = index[FS00B_key_('Tiền SDĐ trước VAT')];
  index.tienthuedattruocvat = index[FS00B_key_('Tiền thuê đất trước VAT')];
  index.htkttruocvat = index[FS00B_key_('HTKT trước VAT')];
  index.chiphibanhangtruocvat = index[FS00B_key_('Chi phí bán hàng trước VAT')];
  index.chiphivanhanhtruocvat = index[FS00B_key_('Chi phí vận hành trước VAT')];
  index.chiphibaotritruocvat = index[FS00B_key_('Chi phí bảo trì trước VAT')];
  index.chiphiduphongtruocvat = index[FS00B_key_('Chi phí dự phòng trước VAT')];
  index.laivayphanbo = FS00B_findHeaderIndex_(index, [
    'Lãi vay phân bổ',
    'Chi phí lãi vay phân bổ',
    'Lãi vay'
  ]);
  index.thuetndn = FS00B_findHeaderIndex_(index, [
    'Thuế TNDN',
    'Thuế TNDN phải nộp',
    'Thuế TNDN tạm tính'
  ]);
  index.lnst = FS00B_findHeaderIndex_(index, [
    'LNST',
    'Lợi nhuận sau thuế'
  ]);
  index.giainganvay = index[FS00B_key_('Giải ngân vay')];
  index.tragoc = index[FS00B_key_('Trả gốc')];

  return { sheet, values, header: values[0], rows: values.slice(1), index };
}

function FS00B_findHeaderIndex_(index, aliases) {
  for (const alias of aliases) {
    const key = FS00B_key_(alias);
    if (index[key] != null) return index[key];
  }
  return null;
}

function FS00B_require_(table, keys, sheetName) {
  const missing = keys.filter(key => {
    const normalized = FS00B_key_(key);
    const direct = table.index[key];
    const normalizedValue = table.index[normalized];
    return direct == null && normalizedValue == null;
  });
  if (missing.length) {
    throw new Error('Thiếu cột tại ' + sheetName + ': ' + missing.join(', '));
  }
}

function FS00B_getProducts_(revenue) {
  const seen = {};
  const products = [];
  revenue.rows.forEach(row => {
    const code = String(row[revenue.index.masp] || '').trim().toUpperCase();
    if (!code || seen[code]) return;
    seen[code] = true;
    products.push({
      code,
      name: String(row[revenue.index.tensanpham] || code).trim(),
      group: String(row[revenue.index.nhom] || '').trim()
    });
  });
  return products;
}

function FS00B_rowsByCode_(table, code) {
  return table.rows.filter(row =>
    String(row[table.index.masp] || '').trim().toUpperCase() === code
  );
}

function FS00B_indexRowsByMonth_(rows, table, monthKey) {
  const map = {};
  const idx = table.index[monthKey] != null ? table.index[monthKey] : table.index[FS00B_key_(monthKey)];
  rows.forEach(row => {
    const monthNo = Math.round(FS00B_num_(row[idx]));
    if (monthNo > 0) map[monthNo] = row;
  });
  return map;
}

function FS00B_buildProjectFinance_(cash) {
  const months = [];
  const dates = [];
  const draw = [];
  const principal = [];

  cash.rows.forEach(row => {
    const monthNo = Math.round(FS00B_num_(row[cash.index.thangso]));
    if (!monthNo) return;
    months.push(monthNo);
    dates.push(FS00B_asDate_(row[cash.index.thang]));
    draw.push(FS00B_num_(row[cash.index.giainganvay]));
    principal.push(FS00B_num_(row[cash.index.tragoc]));
  });

  return { months, dates, draw, principal };
}

function FS00B_buildAllProductCostMap_(cost, products) {
  const result = {};
  products.forEach(p => result[p.code] = {});
  cost.rows.forEach(row => {
    const code = String(row[cost.index.masp] || '').trim().toUpperCase();
    if (!result[code]) return;
    const monthNo = Math.round(FS00B_num_(row[cost.index.thangso]));
    result[code][monthNo] = FS00B_num_(row[cost.index.tongchisauvat]);
  });
  return result;
}

function FS00B_sumCostByMonth_(allProductCosts) {
  const totals = {};
  Object.keys(allProductCosts).forEach(code => {
    Object.keys(allProductCosts[code]).forEach(monthNo => {
      totals[monthNo] = FS00B_num_(totals[monthNo]) + FS00B_num_(allProductCosts[code][monthNo]);
    });
  });
  return totals;
}

function FS00B_globalCostShares_(allProductCosts, products) {
  const sums = {};
  let total = 0;
  products.forEach(p => {
    const value = Object.values(allProductCosts[p.code] || {})
      .reduce((a, b) => a + FS00B_num_(b), 0);
    sums[p.code] = value;
    total += value;
  });
  const shares = {};
  products.forEach(p => shares[p.code] = total > 0 ? sums[p.code] / total : 0);
  return shares;
}

function FS00B_readDiscountRates_(summary) {
  const lastRow = summary.getLastRow();
  const values = summary.getRange(1, 1, lastRow, Math.min(8, summary.getLastColumn())).getValues();

  let wacc = null;
  let costOfEquity = null;

  values.forEach(row => {
    const label = FS00B_key_(row[1]);
    if (label === FS00B_key_('WACC')) {
      for (let c = row.length - 1; c >= 0; c--) {
        const rate = FS00B_rate_(row[c]);
        if (rate > 0) {
          wacc = rate;
          break;
        }
      }
    }
    if (label === FS00B_key_('Vốn chủ sở hữu')) {
      for (let c = row.length - 1; c >= 0; c--) {
        const rate = FS00B_rate_(row[c]);
        if (rate > 0 && rate < 1) {
          costOfEquity = rate;
          break;
        }
      }
    }
  });

  if (!(wacc > 0)) throw new Error('Không đọc được WACC tại Sheet 00.');
  if (!(costOfEquity > 0)) costOfEquity = wacc;

  return { wacc, costOfEquity };
}

function FS00B_summaryRows_(summary) {
  const lastRow = summary.getLastRow();
  const values = summary.getRange(1, 2, lastRow, 1).getDisplayValues();
  const map = {};
  values.forEach((row, i) => {
    const key = FS00B_key_(row[0]);
    if (key && map[key] == null) map[key] = i + 1;
  });
  return map;
}

function FS00B_setByAliases_(sheet, rowMap, aliases, col, value) {
  for (const alias of aliases) {
    const row = rowMap[FS00B_key_(alias)];
    if (row) {
      sheet.getRange(row, col).setValue(FS00B_finiteOrBlank_(value));
      return true;
    }
  }
  return false;
}

function FS00B_formatPercentRows_(sheet, rowMap, aliases, firstCol, count) {
  for (const alias of aliases) {
    const row = rowMap[FS00B_key_(alias)];
    if (row) {
      sheet.getRange(row, firstCol, 1, count).setNumberFormat('0.00%');
      return;
    }
  }
}

function FS00B_formatPaybackRows_(sheet, rowMap, aliases, firstCol, count) {
  for (const alias of aliases) {
    const row = rowMap[FS00B_key_(alias)];
    if (row) {
      sheet.getRange(row, firstCol, 1, count).setNumberFormat('0.00');
      return;
    }
  }
}

function FS00B_lastSectionRow_(summary) {
  const values = summary.getRange(1, 2, summary.getLastRow(), 1).getDisplayValues();
  let start = 0;
  let last = 0;
  values.forEach((row, i) => {
    const text = String(row[0] || '').trim();
    if (text.indexOf('III.') === 0) start = i + 1;
    if (start && text) last = i + 1;
  });
  return last;
}

function FS00B_xnpv_(rate, cashFlows, dates) {
  if (!(rate > -1) || !cashFlows.length || cashFlows.length !== dates.length) return 0;
  const d0 = dates[0];
  return cashFlows.reduce((sum, cf, i) => {
    const years = (dates[i].getTime() - d0.getTime()) / 86400000 / 365;
    return sum + FS00B_num_(cf) / Math.pow(1 + rate, years);
  }, 0);
}

function FS00B_xirr_(cashFlows, dates) {
  const hasPositive = cashFlows.some(v => FS00B_num_(v) > 0);
  const hasNegative = cashFlows.some(v => FS00B_num_(v) < 0);
  if (!hasPositive || !hasNegative || cashFlows.length !== dates.length) return 0;

  const f = rate => FS00B_xnpv_(rate, cashFlows, dates);
  let low = -0.9999;
  let high = 1;
  let fLow = f(low);
  let fHigh = f(high);

  let expand = 0;
  while (fLow * fHigh > 0 && expand < 60) {
    high = high * 2 + 0.5;
    fHigh = f(high);
    expand++;
  }
  if (fLow * fHigh > 0) return 0;

  for (let i = 0; i < 200; i++) {
    const mid = (low + high) / 2;
    const fMid = f(mid);
    if (Math.abs(fMid) < 0.01) return mid;
    if (fLow * fMid <= 0) {
      high = mid;
      fHigh = fMid;
    } else {
      low = mid;
      fLow = fMid;
    }
  }
  return (low + high) / 2;
}

function FS00B_sustainablePayback_(cashFlows) {
  const cumulative = [];
  let running = 0;
  let lastNegative = -1;

  cashFlows.forEach((value, i) => {
    running += FS00B_num_(value);
    cumulative.push(running);
    if (running < -FS00B_CFG.tolerance) lastNegative = i;
  });

  if (lastNegative < 0) return 0;
  const recovery = lastNegative + 1;
  if (recovery >= cashFlows.length) return 0;

  for (let i = recovery; i < cumulative.length; i++) {
    if (cumulative[i] < -FS00B_CFG.tolerance) return 0;
  }

  const previous = cumulative[lastNegative];
  const recoveryFlow = FS00B_num_(cashFlows[recovery]);
  if (recoveryFlow <= 0) return recovery;
  return lastNegative + 1 + Math.min(1, Math.max(0, Math.abs(previous) / recoveryFlow));
}

function FS00B_asDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (typeof value === 'number') {
    return new Date(Date.UTC(1899, 11, 30) + Math.round(value) * 86400000);
  }
  const date = new Date(value);
  if (isNaN(date.getTime())) throw new Error('Giá trị tháng không hợp lệ: ' + value);
  return date;
}

function FS00B_rate_(value) {
  const n = FS00B_num_(value);
  return Math.abs(n) > 1 ? n / 100 : n;
}

function FS00B_num_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  if (value == null || value === '') return 0;
  const text = String(value).trim().replace(/\s/g, '');
  if (!text) return 0;
  const normalized = text.indexOf(',') >= 0 && text.indexOf('.') >= 0
    ? text.replace(/\./g, '').replace(',', '.')
    : text.replace(',', '.');
  const n = Number(normalized.replace('%', ''));
  return isFinite(n) ? n : 0;
}

function FS00B_key_(value) {
  return String(value == null ? '' : value)
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/[^a-z0-9]/g, '');
}

function FS00B_finiteOrBlank_(value) {
  return typeof value === 'number' && isFinite(value) ? value : '';
}
