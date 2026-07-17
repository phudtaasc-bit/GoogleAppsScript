const FS00G_CFG = Object.freeze({
  DEBUG_SHEET: 'DEBUG_IRR',
  BILLION: 1e9,
  TOLERANCE: 1
});

/**
 * Tạo bảng kiểm tra dòng tiền từng sản phẩm.
 * Không sửa các sheet lõi.
 * Dùng lại đúng dữ liệu và logic của module 00F.
 */
function FS00G_taoBangDebugIRR() {
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
  const products = FS00F_products_(revenue);
  const rates = FS00F_discountRates_(summary);

  if (!months) throw new Error('Số tháng mô hình phải lớn hơn 0.');
  if (!products.length) throw new Error('Không có sản phẩm để kiểm tra.');

  const rows = [];
  const diagnostics = [];

  products.forEach(product => {
    const result = FS00G_buildProductSeries_({
      product,
      months,
      loanRatio,
      rates,
      revenue,
      cost,
      tax
    });

    result.rows.forEach(row => rows.push(row));
    diagnostics.push(result.diagnostic);
  });

  FS00G_writeDebugSheet_(ss, diagnostics, rows);
  ss.toast('Đã tạo sheet DEBUG_IRR.', 'Kiểm tra IRR', 6);
}

function FS00G_buildProductSeries_(ctx) {
  const code = ctx.product.code;
  const revByMonth = FS00F_rowsByMonth_(ctx.revenue, code);
  const costByMonth = FS00F_rowsByMonth_(ctx.cost, code);
  const taxByMonth = FS00F_rowsByMonth_(ctx.tax, code);

  let openingCash = 0;
  let openingDebt = 0;
  let openingVatCredit = 0;
  let cumulativeFcff = 0;
  let cumulativeFcfe = 0;

  const dates = [];
  const fcffSeries = [];
  const fcfeSeries = [];
  const rows = [];

  for (let monthNo = 1; monthNo <= ctx.months; monthNo++) {
    const r = revByMonth[monthNo] || null;
    const c = costByMonth[monthNo] || null;
    const t = taxByMonth[monthNo] || null;

    const date = FS00F_date_(r, c, t, ctx.revenue, ctx.cost, ctx.tax, monthNo);
    const customerCash = r ? FS00F_num_(r[ctx.revenue.alias.customerCash]) : 0;
    const vatOut = r ? FS00F_num_(r[ctx.revenue.alias.vatOut]) : 0;
    const vatIn = c ? FS00F_num_(c[ctx.cost.alias.vatIn]) : 0;
    const costAfterVat = c ? FS00F_num_(c[ctx.cost.alias.costAfterVat]) : 0;
    const cit = t ? FS00F_num_(t[ctx.tax.alias.cit]) : 0;
    const fixedInterest = t ? FS00F_num_(t[ctx.tax.alias.capitalizedInterest]) : 0;

    const vatPayable = Math.max(0, vatOut - openingVatCredit - vatIn);
    const closingVatCredit = Math.max(0, openingVatCredit + vatIn - vatOut);

    const fcff = customerCash - costAfterVat - vatPayable - cit;
    const cashBeforeFinancing = openingCash + fcff - fixedInterest;

    let fundingNeed = 0;
    let equityContribution = 0;
    let loanDrawdown = 0;
    let principalRepayment = 0;
    let closingDebt = openingDebt;
    let closingCash = 0;

    if (cashBeforeFinancing < 0) {
      fundingNeed = -cashBeforeFinancing;
      loanDrawdown = fundingNeed * ctx.loanRatio;
      equityContribution = fundingNeed - loanDrawdown;
      closingDebt = openingDebt + loanDrawdown;
    } else {
      principalRepayment = Math.min(openingDebt, cashBeforeFinancing);
      closingDebt = Math.max(0, openingDebt - principalRepayment);
      closingCash = Math.max(0, cashBeforeFinancing - principalRepayment);
    }

    const fcfe = closingCash - openingCash - equityContribution;
    cumulativeFcff += fcff;
    cumulativeFcfe += fcfe;

    dates.push(date);
    fcffSeries.push(fcff);
    fcfeSeries.push(fcfe);

    rows.push([
      ctx.product.code,
      ctx.product.name,
      ctx.product.group,
      monthNo,
      date,
      customerCash,
      costAfterVat,
      vatOut,
      vatIn,
      vatPayable,
      cit,
      fixedInterest,
      fcff,
      cumulativeFcff,
      fundingNeed,
      equityContribution,
      loanDrawdown,
      principalRepayment,
      closingDebt,
      fcfe,
      cumulativeFcfe
    ]);

    openingCash = closingCash;
    openingDebt = closingDebt;
    openingVatCredit = closingVatCredit;
  }

  const positiveFcff = fcffSeries.filter(value => value > FS00G_CFG.TOLERANCE).length;
  const negativeFcff = fcffSeries.filter(value => value < -FS00G_CFG.TOLERANCE).length;
  const positiveFcfe = fcfeSeries.filter(value => value > FS00G_CFG.TOLERANCE).length;
  const negativeFcfe = fcfeSeries.filter(value => value < -FS00G_CFG.TOLERANCE).length;

  return {
    rows,
    diagnostic: [
      ctx.product.code,
      ctx.product.name,
      positiveFcff,
      negativeFcff,
      Math.min.apply(null, fcffSeries) / FS00G_CFG.BILLION,
      Math.max.apply(null, fcffSeries) / FS00G_CFG.BILLION,
      fcffSeries.reduce((sum, value) => sum + value, 0) / FS00G_CFG.BILLION,
      FS00F_xnpv_(ctx.rates.wacc, fcffSeries, dates) / FS00G_CFG.BILLION,
      FS00F_xirr_(fcffSeries, dates),
      positiveFcfe,
      negativeFcfe,
      Math.min.apply(null, fcfeSeries) / FS00G_CFG.BILLION,
      Math.max.apply(null, fcfeSeries) / FS00G_CFG.BILLION,
      fcfeSeries.reduce((sum, value) => sum + value, 0) / FS00G_CFG.BILLION,
      FS00F_xnpv_(ctx.rates.costOfEquity, fcfeSeries, dates) / FS00G_CFG.BILLION,
      FS00F_xirr_(fcfeSeries, dates),
      FS00G_irrStatus_(positiveFcff, negativeFcff, FS00F_xirr_(fcffSeries, dates))
    ]
  };
}

function FS00G_irrStatus_(positiveCount, negativeCount, irr) {
  if (!positiveCount) return 'Không có FCFF dương';
  if (!negativeCount) return 'Không có FCFF âm';
  if (irr === 0) return 'Có đổi dấu nhưng hàm XIRR chưa tìm được nghiệm';
  return 'Có nghiệm IRR';
}

function FS00G_writeDebugSheet_(ss, diagnostics, rows) {
  let sheet = ss.getSheetByName(FS00G_CFG.DEBUG_SHEET);
  if (!sheet) sheet = ss.insertSheet(FS00G_CFG.DEBUG_SHEET);
  sheet.clear();
  sheet.clearFormats();

  sheet.getRange('A1').setValue('CHẨN ĐOÁN IRR THEO SẢN PHẨM').setFontWeight('bold').setFontSize(14);

  const diagnosticHeaders = [[
    'Mã SP', 'Tên sản phẩm',
    'Số kỳ FCFF dương', 'Số kỳ FCFF âm', 'FCFF nhỏ nhất (tỷ)', 'FCFF lớn nhất (tỷ)',
    'Tổng FCFF (tỷ)', 'NPV dự án (tỷ)', 'IRR dự án',
    'Số kỳ FCFE dương', 'Số kỳ FCFE âm', 'FCFE nhỏ nhất (tỷ)', 'FCFE lớn nhất (tỷ)',
    'Tổng FCFE (tỷ)', 'NPV vốn CSH (tỷ)', 'IRR vốn CSH', 'Kết luận IRR dự án'
  ]];

  sheet.getRange(3, 1, 1, diagnosticHeaders[0].length).setValues(diagnosticHeaders);
  if (diagnostics.length) sheet.getRange(4, 1, diagnostics.length, diagnosticHeaders[0].length).setValues(diagnostics);

  const detailStartRow = 6 + diagnostics.length;
  sheet.getRange(detailStartRow, 1).setValue('CHI TIẾT DÒNG TIỀN THEO THÁNG').setFontWeight('bold');

  const detailHeaders = [[
    'Mã SP', 'Tên sản phẩm', 'Nhóm', 'Tháng số', 'Tháng',
    'Dòng tiền khách hàng', 'Tổng chi sau VAT', 'VAT đầu ra', 'VAT đầu vào', 'VAT phải nộp',
    'Thuế TNDN', 'Lãi vay vốn hóa phân bổ', 'FCFF', 'FCFF lũy kế',
    'Nhu cầu vốn', 'Vốn góp CSH', 'Giải ngân vay', 'Trả gốc', 'Dư nợ cuối kỳ',
    'FCFE', 'FCFE lũy kế'
  ]];

  sheet.getRange(detailStartRow + 1, 1, 1, detailHeaders[0].length).setValues(detailHeaders);
  if (rows.length) sheet.getRange(detailStartRow + 2, 1, rows.length, detailHeaders[0].length).setValues(rows);

  sheet.setFrozenRows(detailStartRow + 1);
  sheet.getRange(3, 1, 1, diagnosticHeaders[0].length).setFontWeight('bold').setBackground('#d9ead3').setWrap(true);
  sheet.getRange(detailStartRow + 1, 1, 1, detailHeaders[0].length).setFontWeight('bold').setBackground('#d9eaf7').setWrap(true);

  if (diagnostics.length) {
    sheet.getRange(4, 9, diagnostics.length, 1).setNumberFormat('0.00%');
    sheet.getRange(4, 16, diagnostics.length, 1).setNumberFormat('0.00%');
    sheet.getRange(4, 5, diagnostics.length, 4).setNumberFormat('#,##0.0');
    sheet.getRange(4, 12, diagnostics.length, 4).setNumberFormat('#,##0.0');
  }

  if (rows.length) {
    sheet.getRange(detailStartRow + 2, 5, rows.length, 1).setNumberFormat('MM/yyyy');
    sheet.getRange(detailStartRow + 2, 6, rows.length, 16).setNumberFormat('#,##0');
  }

  sheet.autoResizeColumns(1, detailHeaders[0].length);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(17, 260);
}
