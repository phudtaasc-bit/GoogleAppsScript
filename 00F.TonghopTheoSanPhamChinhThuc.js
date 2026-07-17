const FS00F_CFG = Object.freeze({
  SUMMARY: '00. Tổng hợp',
  TECH: '01A. Kỹ thuật',
  REVENUE: '02. Doanh thu',
  COST: '03. Chi phí & Vốn',
  TAX: '03A. Lợi nhuận & Thuế',
  FIRST_PRODUCT_COL: 5,
  BILLION: 1e9,
  TOLERANCE: 1
});

/**
 * PHƯƠNG ÁN CHÍNH THỨC
 * - Không sửa và không chạy lại 02, 03, 03A, 04, 04A.
 * - Lọc dữ liệu theo Mã SP từ 02, 03, 03A.
 * - Dựng lại dòng tiền kiểu Sheet 04 hoàn toàn trong bộ nhớ.
 * - Lãi vay dùng đúng "Lãi vay vốn hóa phân bổ" tại 03A; không hội tụ lại.
 * - Ghi kết quả vào các cột sản phẩm tại Mục III Sheet 00.
 */
function FS00F_phanTichHieuQuaTheoSanPham() {
  const ss = SpreadsheetApp.getActive();
  const summary = FS00F_sheet_(ss, FS00F_CFG.SUMMARY);
  const tech = FS00F_sheet_(ss, FS00F_CFG.TECH);
  const revenue = FS00F_readTable_(FS00F_sheet_(ss, FS00F_CFG.REVENUE));
  const cost = FS00F_readTable_(FS00F_sheet_(ss, FS00F_CFG.COST));
  const tax = FS00F_readTable_(FS00F_sheet_(ss, FS00F_CFG.TAX));

  FS00F_validateRevenue_(revenue);
  FS00F_validateCost_(cost);
  FS00F_validateTax_(tax);

  const months = Math.max(0, Math.round(FS00F_num_(FS00F_readInfo_(tech, 'Số tháng mô hình'))));
  const loanRatio = FS00F_rate_(FS00F_readInfo_(tech, 'Tỷ lệ vốn vay'));
  if (!months) throw new Error('Số tháng mô hình phải lớn hơn 0.');
  if (loanRatio < 0 || loanRatio > 1) throw new Error('Tỷ lệ vốn vay không hợp lệ.');

  const products = FS00F_products_(revenue);
  if (!products.length) throw new Error('Không có sản phẩm để phân tích.');

  const rates = FS00F_discountRates_(summary);
  const layout = FS00C_chuanBiCotSanPham_(summary, products);

  const results = products.map(product => FS00F_analyzeProduct_({
    product,
    months,
    loanRatio,
    rates,
    revenue,
    cost,
    tax
  }));

  FS00F_writeSummary_(summary, layout, products, results);
  SpreadsheetApp.flush();
  ss.toast('Đã cập nhật phân tích hiệu quả theo từng sản phẩm.', 'FS sản phẩm', 6);
  return results;
}

function FS00F_analyzeProduct_(ctx) {
  const code = ctx.product.code;
  const revByMonth = FS00F_rowsByMonth_(ctx.revenue, code);
  const costByMonth = FS00F_rowsByMonth_(ctx.cost, code);
  const taxByMonth = FS00F_rowsByMonth_(ctx.tax, code);

  const dates = [];
  const fcffSeries = [];
  const fcfeSeries = [];

  let openingCash = 0;
  let openingDebt = 0;
  let openingVatCredit = 0;
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

  for (let monthNo = 1; monthNo <= ctx.months; monthNo++) {
    const r = revByMonth[monthNo] || null;
    const c = costByMonth[monthNo] || null;
    const t = taxByMonth[monthNo] || null;

    const date = FS00F_date_(r, c, t, ctx.revenue, ctx.cost, ctx.tax, monthNo);
    const customerCash = r ? FS00F_num_(r[ctx.revenue.alias.customerCash]) : 0;
    const vatOut = r ? FS00F_num_(r[ctx.revenue.alias.vatOut]) : 0;

    const costBeforeVat = c ? FS00F_num_(c[ctx.cost.alias.costBeforeVat]) : 0;
    const vatIn = c ? FS00F_num_(c[ctx.cost.alias.vatIn]) : 0;
    const costAfterVat = c ? FS00F_num_(c[ctx.cost.alias.costAfterVat]) : 0;

    const cit = t ? FS00F_num_(t[ctx.tax.alias.cit]) : 0;
    const lnst = t ? FS00F_num_(t[ctx.tax.alias.lnst]) : 0;
    const fixedInterest = t ? FS00F_num_(t[ctx.tax.alias.capitalizedInterest]) : 0;

    const vatPayable = Math.max(0, vatOut - openingVatCredit - vatIn);
    const closingVatCredit = Math.max(0, openingVatCredit + vatIn - vatOut);

    const fcff = customerCash - costAfterVat - vatPayable - cit;
    const cashBeforeFinancing = openingCash + fcff - fixedInterest;

    let equityContribution = 0;
    let loanDrawdown = 0;
    let principalRepayment = 0;
    let closingDebt = openingDebt;
    let closingCash = 0;

    if (cashBeforeFinancing < 0) {
      const fundingNeed = -cashBeforeFinancing;
      loanDrawdown = fundingNeed * ctx.loanRatio;
      equityContribution = fundingNeed - loanDrawdown;
      closingDebt = openingDebt + loanDrawdown;
      closingCash = 0;
    } else {
      principalRepayment = Math.min(openingDebt, cashBeforeFinancing);
      closingDebt = Math.max(0, openingDebt - principalRepayment);
      closingCash = Math.max(0, cashBeforeFinancing - principalRepayment);
    }

    const fcfe = closingCash - openingCash - equityContribution;

    dates.push(date);
    fcffSeries.push(fcff);
    fcfeSeries.push(fcfe);

    peakDebt = Math.max(peakDebt, closingDebt);
    sums.revenueWithVat += customerCash;
    sums.totalCostWithVat += costAfterVat + fixedInterest;
    sums.lnst += lnst;
    sums.interest += fixedInterest;
    sums.cit += cit;
    sums.vatPayable += vatPayable;

    if (c) {
      const ratio = costBeforeVat > FS00F_CFG.TOLERANCE ? costAfterVat / costBeforeVat : 1;
      sums.sellingWithVat += FS00F_num_(c[ctx.cost.alias.selling]) * ratio;
      sums.operatingWithVat += FS00F_num_(c[ctx.cost.alias.operating]) * ratio;
      sums.maintenanceWithVat += FS00F_num_(c[ctx.cost.alias.maintenance]) * ratio;

      const coreBeforeVat =
        FS00F_num_(c[ctx.cost.alias.construction]) +
        FS00F_num_(c[ctx.cost.alias.clearance]) +
        FS00F_num_(c[ctx.cost.alias.landUse]) +
        FS00F_num_(c[ctx.cost.alias.landRent]) +
        FS00F_num_(c[ctx.cost.alias.infrastructure]) +
        FS00F_num_(c[ctx.cost.alias.contingency]);
      sums.coreInvestmentWithVat += coreBeforeVat * ratio;
    }

    openingCash = closingCash;
    openingDebt = closingDebt;
    openingVatCredit = closingVatCredit;
  }

  sums.coreInvestmentWithVat += sums.interest;

  return {
    code: ctx.product.code,
    name: ctx.product.name,
    group: ctx.product.group,
    revenueWithVat: sums.revenueWithVat,
    totalCostWithVat: sums.totalCostWithVat,
    coreInvestmentWithVat: sums.coreInvestmentWithVat,
    sellingWithVat: sums.sellingWithVat,
    operatingWithVat: sums.operatingWithVat,
    maintenanceWithVat: sums.maintenanceWithVat,
    lnst: sums.lnst,
    npvProject: FS00F_xnpv_(ctx.rates.wacc, fcffSeries, dates),
    irrProject: FS00F_xirr_(fcffSeries, dates),
    paybackProject: FS00F_payback_(fcffSeries),
    npvEquity: FS00F_xnpv_(ctx.rates.costOfEquity, fcfeSeries, dates),
    irrEquity: FS00F_xirr_(fcfeSeries, dates),
    paybackEquity: FS00F_payback_(fcfeSeries),
    peakDebt,
    totalInterest: sums.interest,
    totalCit: sums.cit,
    totalVatPayable: sums.vatPayable
  };
}

function FS00F_writeSummary_(sheet, layout, products, results) {
  const rows = FS00F_summaryRows_(sheet);
  const byCode = {};
  results.forEach(result => byCode[result.code] = result);

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

  products.forEach((product, index) => {
    const col = layout.firstProductColumn + index;
    const result = byCode[product.code];

    FS00F_set_(sheet, rows, labels.totalRevenue, col, result.revenueWithVat / FS00F_CFG.BILLION);
    FS00F_set_(sheet, rows, labels.totalCost, col, result.totalCostWithVat / FS00F_CFG.BILLION);
    FS00F_set_(sheet, rows, labels.coreInvestment, col, result.coreInvestmentWithVat / FS00F_CFG.BILLION);
    FS00F_set_(sheet, rows, labels.selling, col, result.sellingWithVat / FS00F_CFG.BILLION);
    FS00F_set_(sheet, rows, labels.operating, col, result.operatingWithVat / FS00F_CFG.BILLION);
    FS00F_set_(sheet, rows, labels.maintenance, col, result.maintenanceWithVat / FS00F_CFG.BILLION);
    FS00F_set_(sheet, rows, labels.lnst, col, result.lnst / FS00F_CFG.BILLION);
    FS00F_set_(sheet, rows, labels.npvProject, col, result.npvProject / FS00F_CFG.BILLION);
    FS00F_set_(sheet, rows, labels.irrProject, col, result.irrProject);
    FS00F_set_(sheet, rows, labels.paybackProject, col, result.paybackProject);
    FS00F_set_(sheet, rows, labels.npvEquity, col, result.npvEquity / FS00F_CFG.BILLION);
    FS00F_set_(sheet, rows, labels.irrEquity, col, result.irrEquity);
    FS00F_set_(sheet, rows, labels.paybackEquity, col, result.paybackEquity);
    FS00F_set_(sheet, rows, labels.peakDebt, col, result.peakDebt / FS00F_CFG.BILLION);
    FS00F_set_(sheet, rows, labels.interest, col, result.totalInterest / FS00F_CFG.BILLION);
    FS00F_set_(sheet, rows, labels.cit, col, result.totalCit / FS00F_CFG.BILLION);
    FS00F_set_(sheet, rows, labels.vat, col, result.totalVatPayable / FS00F_CFG.BILLION);

    products.forEach(detail => {
      FS00F_set_(sheet, rows, [
        'Phần ' + detail.name,
        'Phần ' + detail.name + (detail.group ? ' - ' + detail.group : '')
      ], col, detail.code === product.code ? result.revenueWithVat / FS00F_CFG.BILLION : 0);
    });
  });

  FS00F_formatRow_(sheet, rows, labels.irrProject, layout.firstProductColumn, products.length, '0.00%');
  FS00F_formatRow_(sheet, rows, labels.irrEquity, layout.firstProductColumn, products.length, '0.00%');
  FS00F_formatRow_(sheet, rows, labels.paybackProject, layout.firstProductColumn, products.length, '0.00');
  FS00F_formatRow_(sheet, rows, labels.paybackEquity, layout.firstProductColumn, products.length, '0.00');
}

function FS00F_validateRevenue_(table) {
  table.alias = {
    monthNo: FS00F_col_(table, ['Tháng số']),
    date: FS00F_col_(table, ['Tháng']),
    code: FS00F_col_(table, ['Mã SP']),
    name: FS00F_col_(table, ['Tên sản phẩm']),
    group: FS00F_col_(table, ['Nhóm', 'Loại hình']),
    vatOut: FS00F_col_(table, ['VAT đầu ra']),
    customerCash: FS00F_col_(table, ['Dòng tiền khách hàng', 'Dòng tiền thu khách hàng'])
  };
}

function FS00F_validateCost_(table) {
  table.alias = {
    monthNo: FS00F_col_(table, ['Tháng số']),
    date: FS00F_col_(table, ['Tháng']),
    code: FS00F_col_(table, ['Mã SP']),
    costBeforeVat: FS00F_col_(table, ['Tổng chi trước VAT']),
    vatIn: FS00F_col_(table, ['VAT đầu vào']),
    costAfterVat: FS00F_col_(table, ['Tổng chi sau VAT']),
    construction: FS00F_col_(table, ['XD/TB trước VAT']),
    clearance: FS00F_col_(table, ['GPMB trước VAT']),
    landUse: FS00F_col_(table, ['Tiền SDĐ trước VAT']),
    landRent: FS00F_col_(table, ['Tiền thuê đất trước VAT']),
    infrastructure: FS00F_col_(table, ['HTKT trước VAT']),
    selling: FS00F_col_(table, ['Chi phí bán hàng trước VAT']),
    operating: FS00F_col_(table, ['Chi phí vận hành trước VAT']),
    maintenance: FS00F_col_(table, ['Chi phí bảo trì trước VAT']),
    contingency: FS00F_col_(table, ['Chi phí dự phòng trước VAT'])
  };
}

function FS00F_validateTax_(table) {
  table.alias = {
    monthNo: FS00F_col_(table, ['Tháng số']),
    date: FS00F_col_(table, ['Tháng']),
    code: FS00F_col_(table, ['Mã SP']),
    capitalizedInterest: FS00F_col_(table, ['Lãi vay vốn hóa phân bổ']),
    cit: FS00F_col_(table, ['Thuế TNDN']),
    lnst: FS00F_col_(table, ['LNST', 'Lợi nhuận sau thuế'])
  };
}

function FS00F_products_(table) {
  const seen = {};
  const out = [];
  table.rows.forEach(row => {
    const code = String(row[table.alias.code] || '').trim().toUpperCase();
    if (!code || seen[code]) return;
    seen[code] = true;
    out.push({
      code,
      name: String(row[table.alias.name] || code).trim(),
      group: String(row[table.alias.group] || '').trim()
    });
  });
  return out;
}

function FS00F_rowsByMonth_(table, code) {
  const out = {};
  table.rows.forEach(row => {
    if (String(row[table.alias.code] || '').trim().toUpperCase() !== code) return;
    const monthNo = Math.round(FS00F_num_(row[table.alias.monthNo]));
    if (monthNo > 0) out[monthNo] = row;
  });
  return out;
}

function FS00F_date_(r, c, t, revenue, cost, tax, monthNo) {
  const values = [
    r ? r[revenue.alias.date] : null,
    c ? c[cost.alias.date] : null,
    t ? t[tax.alias.date] : null
  ];
  for (const value of values) {
    if (value instanceof Date && !isNaN(value.getTime())) return value;
    if (typeof value === 'number' && value > 0) return new Date(Date.UTC(1899, 11, 30) + Math.round(value) * 86400000);
  }
  return new Date(2000, monthNo - 1, 1);
}

function FS00F_discountRates_(summary) {
  const values = summary.getRange(1, 1, summary.getLastRow(), Math.min(8, summary.getLastColumn())).getValues();
  let wacc = 0;
  let costOfEquity = 0;
  values.forEach(row => {
    const label = FS00F_key_(row[1]);
    if (label === FS00F_key_('WACC')) {
      for (let i = row.length - 1; i >= 0; i--) {
        const value = FS00F_rate_(row[i]);
        if (value > 0) { wacc = value; break; }
      }
    }
    if (label === FS00F_key_('Vốn chủ sở hữu')) {
      for (let i = row.length - 1; i >= 0; i--) {
        const value = FS00F_rate_(row[i]);
        if (value > 0 && value < 1) { costOfEquity = value; break; }
      }
    }
  });
  if (!(wacc > 0)) throw new Error('Không đọc được WACC trên Sheet 00.');
  if (!(costOfEquity > 0)) costOfEquity = wacc;
  return { wacc, costOfEquity };
}

function FS00F_xnpv_(rate, cashFlows, dates) {
  if (!(rate > -1) || !cashFlows.length) return 0;
  const d0 = dates[0];
  return cashFlows.reduce((sum, value, index) => {
    const years = (dates[index].getTime() - d0.getTime()) / 86400000 / 365;
    return sum + FS00F_num_(value) / Math.pow(1 + rate, years);
  }, 0);
}

function FS00F_xirr_(cashFlows, dates) {
  if (!cashFlows.some(v => v > 0) || !cashFlows.some(v => v < 0)) return 0;
  const fn = rate => FS00F_xnpv_(rate, cashFlows, dates);
  let low = -0.9999;
  let high = 1;
  let fLow = fn(low);
  let fHigh = fn(high);
  let count = 0;
  while (fLow * fHigh > 0 && count < 60) {
    high = high * 2 + 0.5;
    fHigh = fn(high);
    count++;
  }
  if (fLow * fHigh > 0) return 0;
  for (let i = 0; i < 200; i++) {
    const mid = (low + high) / 2;
    const fMid = fn(mid);
    if (Math.abs(fMid) < 0.01) return mid;
    if (fLow * fMid <= 0) {
      high = mid;
    } else {
      low = mid;
      fLow = fMid;
    }
  }
  return (low + high) / 2;
}

function FS00F_payback_(cashFlows) {
  let running = 0;
  let lastNegative = -1;
  const cumulative = [];
  cashFlows.forEach((value, index) => {
    running += FS00F_num_(value);
    cumulative.push(running);
    if (running < -FS00F_CFG.TOLERANCE) lastNegative = index;
  });
  if (lastNegative < 0) return 0;
  const recovery = lastNegative + 1;
  if (recovery >= cashFlows.length) return 0;
  for (let i = recovery; i < cumulative.length; i++) if (cumulative[i] < -FS00F_CFG.TOLERANCE) return 0;
  const flow = FS00F_num_(cashFlows[recovery]);
  if (flow <= 0) return recovery;
  return lastNegative + 1 + Math.min(1, Math.max(0, Math.abs(cumulative[lastNegative]) / flow));
}

function FS00F_summaryRows_(sheet) {
  const values = sheet.getRange(1, 2, sheet.getLastRow(), 1).getDisplayValues();
  const out = {};
  values.forEach((row, index) => {
    const key = FS00F_key_(row[0]);
    if (key && out[key] == null) out[key] = index + 1;
  });
  return out;
}

function FS00F_set_(sheet, rows, aliases, col, value) {
  for (const alias of aliases) {
    const row = rows[FS00F_key_(alias)];
    if (row) {
      sheet.getRange(row, col).setValue(isFinite(value) ? value : '');
      return;
    }
  }
}

function FS00F_formatRow_(sheet, rows, aliases, firstCol, count, format) {
  for (const alias of aliases) {
    const row = rows[FS00F_key_(alias)];
    if (row) {
      sheet.getRange(row, firstCol, 1, count).setNumberFormat(format);
      return;
    }
  }
}

function FS00F_readTable_(sheet) {
  const cols = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, cols).getDisplayValues()[0];
  const rows = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, cols).getValues() : [];
  const index = {};
  headers.forEach((header, col) => index[FS00F_key_(header)] = col);
  return { sheet, headers, rows, index, alias: {} };
}

function FS00F_col_(table, aliases) {
  for (const alias of aliases) {
    const col = table.index[FS00F_key_(alias)];
    if (col != null) return col;
  }
  throw new Error('Sheet "' + table.sheet.getName() + '" thiếu cột: ' + aliases.join(' / '));
}

function FS00F_sheet_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Không tìm thấy sheet: ' + name);
  return sheet;
}

function FS00F_readInfo_(sheet, label) {
  const target = FS00F_key_(label);
  const values = sheet.getDataRange().getValues();
  for (let i = 0; i < values.length; i++) if (FS00F_key_(values[i][0]) === target) return values[i][1];
  return '';
}

function FS00F_num_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  const text = String(value == null ? '' : value).trim().replace(/\s/g, '');
  if (!text) return 0;
  const number = Number(text.includes(',') && text.includes('.') ? text.replace(/\./g, '').replace(',', '.') : text.replace(/,/g, ''));
  return isFinite(number) ? number : 0;
}

function FS00F_rate_(value) {
  if (typeof value === 'number') return value > 1 ? value / 100 : value;
  const text = String(value == null ? '' : value).trim();
  if (!text) return 0;
  const number = FS00F_num_(text.replace('%', ''));
  return text.includes('%') || number > 1 ? number / 100 : number;
}

function FS00F_key_(value) {
  return String(value == null ? '' : value)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/²/g, '2')
    .replace(/\^2/g, '2')
    .replace(/m\s*2/g, 'm2')
    .replace(/[^a-z0-9]/g, '');
}
