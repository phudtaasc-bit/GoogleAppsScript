const FS05_CFG = Object.freeze({
  SHEET: '05. Độ nhạy',
  TECH: '01A. Kỹ thuật',
  REVENUE: '02. Doanh thu',
  COST: '03. Chi phí & Vốn',
  PROFIT: '03A. Lợi nhuận & Thuế',
  CASH: '04. Dòng tiền & Tài trợ',
  SUMMARY: '00. Tổng hợp',
  FACTORS: [-0.10, -0.08, -0.05, -0.04, -0.02, 0, 0.02, 0.04, 0.05, 0.08, 0.10],
  DEFAULT_VAT: 0.08,
  COLOR_BASE: '#008000',
  COLOR_CENTER: '#FFFF00',
  COLOR_LOW: '#E6B8B7',
  COLOR_WHITE: '#FFFFFF',
  COLOR_HEADER: '#A6A6A6',
  COLOR_SECTION: '#FFC000',
  COLOR_TITLE: '#1F4E78'
});

function FS05_DoNhay_Fast() {
  return FS05_lapBangDoNhay();
}

function FS05B_DoNhay_CSH() {
  return FS05_lapBangDoNhay();
}

function FS05_lapBangDoNhay() {
  const ss = SpreadsheetApp.getActive();
  const required = [FS05_CFG.TECH, FS05_CFG.REVENUE, FS05_CFG.COST, FS05_CFG.PROFIT, FS05_CFG.CASH, FS05_CFG.SUMMARY];
  required.forEach(name => {
    if (!ss.getSheetByName(name)) throw new Error('Thiếu sheet "' + name + '". Hãy chạy toàn bộ mô hình trước.');
  });

  const model = FS05_readModel_(ss);
  const baseOfficial = FS05_readOfficialBase_(ss.getSheetByName(FS05_CFG.SUMMARY));
  const baseEngine = FS05_runScenario_(model, 1, 1, model.loanRate);

  const sale = FS05_buildMatrix_(model, baseOfficial, baseEngine, 'SALE');
  const rate = FS05_buildMatrix_(model, baseOfficial, baseEngine, 'RATE');
  const rent = FS05_buildMatrix_(model, baseOfficial, baseEngine, 'RENT');

  let sh = ss.getSheetByName(FS05_CFG.SHEET);
  if (!sh) sh = ss.insertSheet(FS05_CFG.SHEET);
  sh.clear();
  sh.clearFormats();

  FS05_layout_(sh);
  FS05_writeBlock_(sh, 4, 'NPV', 'Tăng/giảm giá bán', baseOfficial.npvProject, sale.npvProject, '#,##0.0');
  FS05_writeBlock_(sh, 20, 'IRR', 'Tăng/giảm giá bán', baseOfficial.irrProject, sale.irrProject, '0.0%');
  FS05_writeBlock_(sh, 36, 'NPV VỐN CSH', 'Tăng/giảm lãi suất vay', baseOfficial.npvEquity, rate.npvEquity, '#,##0.0');
  FS05_writeBlock_(sh, 52, 'IRR VỐN CSH', 'Tăng/giảm lãi suất vay', baseOfficial.irrEquity, rate.irrEquity, '0.0%');
  FS05_writeBlock_(sh, 68, 'NPV - GIÁ THUÊ/VỐN ĐẦU TƯ', 'Tăng/giảm giá thuê', baseOfficial.npvProject, rent.npvProject, '#,##0.0');
  FS05_writeBlock_(sh, 84, 'IRR - GIÁ THUÊ/VỐN ĐẦU TƯ', 'Tăng/giảm giá thuê', baseOfficial.irrProject, rent.irrProject, '0.0%');
  FS05_format_(sh);
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Đã hoàn thành 6 bảng độ nhạy trong bộ nhớ, không chạy lại từng sheet theo kịch bản.');
}

function FS05_readModel_(ss) {
  const tech = ss.getSheetByName(FS05_CFG.TECH);
  const revenue = FS05_readTable_(ss.getSheetByName(FS05_CFG.REVENUE));
  const cost = FS05_readTable_(ss.getSheetByName(FS05_CFG.COST));
  const profit = FS05_readTable_(ss.getSheetByName(FS05_CFG.PROFIT));
  const cash = FS05_readTable_(ss.getSheetByName(FS05_CFG.CASH));

  FS05_require_(revenue.index, ['thangso', 'thang', 'masp', 'nhom', 'tongdoanhthutruocvat', 'vatdaura', 'dongtienkhachhang'], FS05_CFG.REVENUE);
  FS05_require_(cost.index, ['thangso', 'masp', 'tongchitruocvat', 'vatdauvao', 'tongchisauvat'], FS05_CFG.COST);
  FS05_require_(profit.index, ['thangso', 'masp', 'tongdoanhthutruocvat', 'laivayvonhoaphanbo'], FS05_CFG.PROFIT);
  FS05_require_(cash.index, ['thangso', 'thang'], FS05_CFG.CASH);

  const products = FS05_readProducts_(tech);
  const productByCode = {};
  products.forEach(p => { productByCode[p.code] = p; });

  const months = Math.max.apply(null, cash.values.map(r => FS05_num_(r[cash.index.thangso])).filter(Boolean));
  const dates = Array(months).fill(null);
  cash.values.forEach(row => {
    const m = FS05_num_(row[cash.index.thangso]);
    if (m >= 1 && m <= months) dates[m - 1] = row[cash.index.thang];
  });
  dates.forEach((d, i) => {
    if (!(d instanceof Date) || isNaN(d.getTime())) throw new Error('Cột Tháng tại Sheet 04 không phải ngày thực, dòng ' + (i + 2) + '.');
  });

  const revenueByKey = FS05_indexRows_(revenue);
  const costByKey = FS05_indexRows_(cost);
  const profitByKey = FS05_indexRows_(profit);
  const keys = Object.keys(revenueByKey);

  const rows = keys.map(key => {
    const rr = revenueByKey[key];
    const cr = costByKey[key] || [];
    const pr = profitByKey[key] || [];
    const code = String(rr[revenue.index.masp] || '').trim().toUpperCase();
    const product = productByCode[code] || { code, group: FS05_group_(rr[revenue.index.nhom]), taxRate: 0.20 };
    const monthNo = FS05_num_(rr[revenue.index.thangso]);
    const group = product.group || FS05_group_(rr[revenue.index.nhom]);

    return {
      key,
      code,
      monthNo,
      group,
      taxRate: product.taxRate,
      revenue: FS05_num_(rr[revenue.index.tongdoanhthutruocvat]),
      vatOut: FS05_num_(rr[revenue.index.vatdaura]),
      customerCash: FS05_num_(rr[revenue.index.dongtienkhachhang]),
      cost: FS05_extractCost_(cr, cost.index),
      accounting: FS05_extractAccounting_(pr, profit.index)
    };
  });

  const interestShares = FS05_interestShares_(rows, months);
  const vatRates = FS05_readVatRates_(tech);

  return {
    months,
    dates,
    rows,
    interestShares,
    vatRates,
    loanRatio: FS05_rate_(FS05_readInfoValue_(tech, 'Tỷ lệ vốn vay')),
    loanRate: FS05_rate_(FS05_readInfoValue_(tech, 'Lãi suất vay năm')),
    equityRate: FS05_rate_(FS05_readInfoValue_(tech, 'Tỷ suất chiết khấu'))
  };
}

function FS05_buildMatrix_(model, official, engineBase, mode) {
  const out = {
    npvProject: FS05_emptyMatrix_(), irrProject: FS05_emptyMatrix_(),
    npvEquity: FS05_emptyMatrix_(), irrEquity: FS05_emptyMatrix_()
  };

  FS05_CFG.FACTORS.forEach((costChange, r) => {
    FS05_CFG.FACTORS.forEach((horizontalChange, c) => {
      let saleFactor = 1;
      let rentFactor = 1;
      let loanRate = model.loanRate;
      if (mode === 'SALE') saleFactor = 1 + horizontalChange;
      if (mode === 'RENT') rentFactor = 1 + horizontalChange;
      if (mode === 'RATE') loanRate = Math.max(0, model.loanRate + horizontalChange);

      const result = FS05_runScenario_(model, saleFactor, rentFactor, loanRate, 1 + costChange);
      out.npvProject[r][c] = official.npvProject + (result.npvProject - engineBase.npvProject) / 1e9;
      out.irrProject[r][c] = official.irrProject + (result.irrProject - engineBase.irrProject);
      out.npvEquity[r][c] = official.npvEquity + (result.npvEquity - engineBase.npvEquity) / 1e9;
      out.irrEquity[r][c] = official.irrEquity + (result.irrEquity - engineBase.irrEquity);

      if (costChange === 0 && horizontalChange === 0) {
        out.npvProject[r][c] = official.npvProject;
        out.irrProject[r][c] = official.irrProject;
        out.npvEquity[r][c] = official.npvEquity;
        out.irrEquity[r][c] = official.irrEquity;
      }
    });
  });
  return out;
}

function FS05_runScenario_(model, saleFactor, rentFactor, annualLoanRate, costFactor) {
  costFactor = costFactor == null ? 1 : costFactor;
  let taxByMonth = Array(model.months).fill(0);
  let financing = null;

  for (let iteration = 0; iteration < 20; iteration++) {
    const monthly = FS05_buildMonthlyCash_(model, saleFactor, rentFactor, costFactor);
    financing = FS05_finance_(model, monthly, taxByMonth, annualLoanRate);
    const nextTax = FS05_buildTax_(model, saleFactor, rentFactor, costFactor, financing.interestByMonth);
    const delta = nextTax.reduce((s, v, i) => s + Math.abs(v - taxByMonth[i]), 0);
    taxByMonth = nextTax;
    if (delta <= 1) break;
  }

  const monthly = FS05_buildMonthlyCash_(model, saleFactor, rentFactor, costFactor);
  financing = FS05_finance_(model, monthly, taxByMonth, annualLoanRate);

  const totalInvestment = monthly.reduce((s, x) => s + x.costAfterVat, 0) + financing.interestByMonth.reduce((s, v) => s + v, 0);
  const equity = financing.equity.reduce((s, v) => s + v, 0);
  const loan = financing.loan.reduce((s, v) => s + v, 0);
  const wacc = totalInvestment > 0
    ? (equity / totalInvestment) * model.equityRate + (loan / totalInvestment) * annualLoanRate
    : model.equityRate;

  return {
    npvProject: FS05_xnpv_(wacc, financing.fcff, model.dates),
    irrProject: FS05_xirr_(financing.fcff, model.dates, 0.10),
    npvEquity: FS05_xnpv_(model.equityRate, financing.fcfe, model.dates),
    irrEquity: FS05_xirr_(financing.fcfe, model.dates, 0.10)
  };
}

function FS05_buildMonthlyCash_(model, saleFactor, rentFactor, costFactor) {
  const out = Array.from({ length: model.months }, () => ({ customerCash: 0, vatOut: 0, vatIn: 0, costBeforeVat: 0, costAfterVat: 0 }));

  model.rows.forEach(row => {
    const rf = row.group === 'Bán' ? saleFactor : rentFactor;
    const m = row.monthNo - 1;
    const scaled = FS05_scaleCost_(row.cost, row.group, rf, rentFactor, costFactor);
    const vatIn = FS05_costVat_(scaled, row.code, model.vatRates);
    out[m].customerCash += row.customerCash * rf;
    out[m].vatOut += row.vatOut * rf;
    out[m].costBeforeVat += scaled.total;
    out[m].vatIn += vatIn;
    out[m].costAfterVat += scaled.total + vatIn;
  });
  return out;
}

function FS05_buildTax_(model, saleFactor, rentFactor, costFactor, interestByMonth) {
  const tax = Array(model.months).fill(0);
  model.rows.forEach(row => {
    const rf = row.group === 'Bán' ? saleFactor : rentFactor;
    const a = row.accounting;
    const nonInterest =
      (a.baseCost + a.land) * costFactor +
      a.selling * rf +
      (a.operating + a.maintenance) * rentFactor;
    const interest = (interestByMonth[row.monthNo - 1] || 0) * (model.interestShares[row.key] || 0);
    const taxable = Math.max(0, row.revenue * rf - nonInterest - interest);
    tax[row.monthNo - 1] += taxable * row.taxRate;
  });
  return tax;
}

function FS05_finance_(model, monthly, taxByMonth, annualLoanRate) {
  const monthlyRate = Math.pow(1 + annualLoanRate, 1 / 12) - 1;
  const fcff = [], fcfe = [], equity = [], loan = [], interestByMonth = [];
  let openingCash = 0;
  let openingDebt = 0;
  let vatCredit = 0;

  for (let i = 0; i < model.months; i++) {
    const x = monthly[i];
    const vatPayable = Math.max(0, x.vatOut - vatCredit - x.vatIn);
    vatCredit = Math.max(0, vatCredit + x.vatIn - x.vatOut);
    const projectCash = x.customerCash - x.costAfterVat - vatPayable - (taxByMonth[i] || 0);
    const interest = openingDebt * monthlyRate;
    const beforeFinancing = openingCash + projectCash - interest;

    let equityContribution = 0;
    let loanDraw = 0;
    let repayment = 0;
    let closingDebt = openingDebt;
    let closingCash = 0;

    if (beforeFinancing < 0) {
      const need = -beforeFinancing;
      loanDraw = need * model.loanRatio;
      equityContribution = need - loanDraw;
      closingDebt += loanDraw;
    } else {
      repayment = Math.min(openingDebt, beforeFinancing);
      closingDebt = Math.max(0, openingDebt - repayment);
      closingCash = Math.max(0, beforeFinancing - repayment);
    }

    fcff.push(projectCash);
    fcfe.push(closingCash - openingCash - equityContribution);
    equity.push(equityContribution);
    loan.push(loanDraw);
    interestByMonth.push(interest);
    openingCash = closingCash;
    openingDebt = closingDebt;
  }
  return { fcff, fcfe, equity, loan, interestByMonth };
}

function FS05_scaleCost_(c, group, revenueFactor, rentFactor, costFactor) {
  const x = {
    xd: c.xd * costFactor,
    gpmb: c.gpmb * costFactor,
    htkt: c.htkt * costFactor,
    sdd: c.sdd * costFactor,
    rentLand: c.rentLand * costFactor,
    contingency: c.contingency * costFactor,
    selling: c.selling * revenueFactor,
    operating: c.operating * rentFactor,
    maintenance: c.maintenance * rentFactor
  };
  x.total = x.xd + x.gpmb + x.htkt + x.sdd + x.rentLand + x.contingency + x.selling + x.operating + x.maintenance;
  return x;
}

function FS05_costVat_(c, code, rates) {
  const landRate = code === 'LK' ? rates.sdd : rates.rentLand;
  return c.xd * rates.xd + c.gpmb * rates.gpmb + c.htkt * rates.htkt +
    c.sdd * rates.sdd + c.rentLand * landRate + c.contingency * rates.contingency +
    c.selling * rates.selling + c.operating * rates.operating + c.maintenance * rates.maintenance;
}

function FS05_extractCost_(row, index) {
  const v = key => index[key] == null ? 0 : FS05_num_(row[index[key]]);
  return {
    xd: v('xdtbtruocvat'), gpmb: v('gpmbtruocvat'), htkt: v('htkttruocvat'),
    sdd: v('tiensddtruocvat'), rentLand: v('tienthuedattruocvat'),
    selling: v('chiphibanhangtruocvat'), operating: v('chiphivanhanhtruocvat'),
    maintenance: v('chiphibaotritruocvat'), contingency: v('chiphiduphongtruocvat')
  };
}

function FS05_extractAccounting_(row, index) {
  const v = key => index[key] == null ? 0 : FS05_num_(row[index[key]]);
  return {
    baseCost: v('giavonxdgpmbhtktduphong'),
    land: v('tiensddphanbo') + v('tienthuedatphanbo'),
    selling: v('chiphibanhang'), operating: v('chiphivanhanh'), maintenance: v('chiphibaotri'),
    interest: v('laivayvonhoaphanbo')
  };
}

function FS05_interestShares_(rows, months) {
  const totals = Array(months).fill(0);
  rows.forEach(r => { totals[r.monthNo - 1] += Math.max(0, r.accounting.interest); });
  const shares = {};
  rows.forEach(r => {
    const total = totals[r.monthNo - 1];
    shares[r.key] = total > 0 ? Math.max(0, r.accounting.interest) / total : 0;
  });
  return shares;
}

function FS05_readProducts_(tech) {
  const block = FS05_findBlock_(tech, 'SAN_PHAM');
  return tech.getRange(block.startRow, 1, block.rowCount, 14).getValues().map(row => ({
    code: String(row[0] || '').trim().toUpperCase(),
    group: FS05_group_(row[2]),
    taxRate: FS05_rate_(row[8])
  })).filter(p => p.code);
}

function FS05_readVatRates_(tech) {
  const block = FS05_findBlock_(tech, 'CHI_PHI_CHUNG');
  const rows = tech.getRange(block.startRow, 1, block.rowCount, 6).getValues();
  const map = {};
  rows.forEach(r => { map[FS05_key_(r[0])] = FS05_rate_(r[2]); });
  const get = (labels, fallback) => {
    for (const label of labels) if (map[FS05_key_(label)] != null) return map[FS05_key_(label)];
    return fallback;
  };
  return {
    xd: get(['Chi phí XD/TB/khác', 'Chi phí XD/TB'], FS05_CFG.DEFAULT_VAT),
    gpmb: get(['Chi phí GPMB'], 0),
    htkt: get(['Chi phí HTKT'], FS05_CFG.DEFAULT_VAT),
    sdd: get(['Tiền SDĐ liền kề'], 0),
    rentLand: get(['Tiền thuê đất chung cư', 'Tiền thuê đất TMDV', 'Tiền thuê đất Chợ'], 0),
    selling: get(['Chi phí bán hàng'], FS05_CFG.DEFAULT_VAT),
    contingency: get(['Chi phí dự phòng'], FS05_CFG.DEFAULT_VAT),
    operating: FS05_CFG.DEFAULT_VAT,
    maintenance: FS05_CFG.DEFAULT_VAT
  };
}

function FS05_readOfficialBase_(summary) {
  const get = label => {
    const row = FS05_findSummaryRow_(summary, label);
    if (!row) throw new Error('Không tìm thấy chỉ tiêu "' + label + '" trên Sheet 00.');
    return FS05_num_(summary.getRange(row, 4).getValue());
  };
  return {
    npvProject: get('NPV dự án'), irrProject: get('IRR dự án'),
    npvEquity: get('NPV vốn CSH'), irrEquity: get('IRR vốn CSH')
  };
}

function FS05_xnpv_(rate, values, dates) {
  if (rate <= -1 || !values.length) return 0;
  const d0 = dates[0].getTime();
  return values.reduce((sum, value, i) => {
    const years = (dates[i].getTime() - d0) / 31536000000;
    return sum + value / Math.pow(1 + rate, years);
  }, 0);
}

function FS05_xirr_(values, dates, guess) {
  if (!values.some(v => v > 0) || !values.some(v => v < 0)) return 0;
  let rate = guess == null ? 0.1 : guess;
  const d0 = dates[0].getTime();
  for (let k = 0; k < 80; k++) {
    let f = 0, df = 0;
    for (let i = 0; i < values.length; i++) {
      const t = (dates[i].getTime() - d0) / 31536000000;
      const den = Math.pow(1 + rate, t);
      f += values[i] / den;
      if (t) df += -t * values[i] / (den * (1 + rate));
    }
    if (!isFinite(f) || !isFinite(df) || Math.abs(df) < 1e-12) break;
    const next = rate - f / df;
    if (!isFinite(next) || next <= -0.999999 || next > 100) break;
    if (Math.abs(next - rate) < 1e-10) return next;
    rate = next;
  }
  return FS05_xirrBisection_(values, dates);
}

function FS05_xirrBisection_(values, dates) {
  let prevRate = -0.99;
  let prevValue = FS05_xnpv_(prevRate, values, dates);
  const candidates = [];
  for (let i = 1; i <= 400; i++) {
    const rate = -0.99 + i * (10.99 / 400);
    const value = FS05_xnpv_(rate, values, dates);
    if (prevValue === 0 || value === 0 || prevValue * value < 0) candidates.push([prevRate, rate]);
    prevRate = rate;
    prevValue = value;
  }
  if (!candidates.length) return 0;
  let pair = candidates.reduce((best, p) => Math.abs((p[0] + p[1]) / 2 - 0.1) < Math.abs((best[0] + best[1]) / 2 - 0.1) ? p : best, candidates[0]);
  let lo = pair[0], hi = pair[1], flo = FS05_xnpv_(lo, values, dates);
  for (let i = 0; i < 100; i++) {
    const mid = (lo + hi) / 2;
    const fm = FS05_xnpv_(mid, values, dates);
    if (Math.abs(fm) < 0.01) return mid;
    if (flo * fm <= 0) hi = mid;
    else { lo = mid; flo = fm; }
  }
  return (lo + hi) / 2;
}

function FS05_layout_(sh) {
  sh.getRange('A1:M1').merge().setValue('05. BẢNG PHÂN TÍCH ĐỘ NHẠY');
  sh.getRange('A2:M2').merge().setValue('Màu hồng: kết quả thấp hơn Base. Màu trắng: kết quả lớn hơn hoặc bằng Base.');
}

function FS05_writeBlock_(sh, startRow, title, horizontalTitle, base, matrix, format) {
  const n = FS05_CFG.FACTORS.length;
  const centerRow = startRow + 8;
  const centerCol = 8;
  sh.getRange(startRow, 1, 1, 13).setBackground(FS05_CFG.COLOR_SECTION).setFontWeight('bold');
  sh.getRange(startRow, 2).setValue(title).setFontWeight('bold').setFontSize(12);
  sh.getRange(startRow + 1, 3, 1, n).merge().setValue(horizontalTitle).setFontWeight('bold');
  sh.getRange(startRow + 2, 2).setValue(base).setNumberFormat(format).setBackground(FS05_CFG.COLOR_BASE).setFontColor('#FFFFFF').setFontWeight('bold');
  sh.getRange(startRow + 2, 3, 1, n).setValues([FS05_CFG.FACTORS]).setNumberFormat('0%').setBackground(FS05_CFG.COLOR_HEADER).setFontColor('#FFFFFF').setFontWeight('bold');
  sh.getRange(startRow + 3, 1, n, 1).merge().setValue('Tăng/giảm\nvốn đầu tư').setFontWeight('bold').setWrap(true);
  sh.getRange(startRow + 3, 2, n, 1).setValues(FS05_CFG.FACTORS.map(x => [x])).setNumberFormat('0%').setFontWeight('bold');
  const data = sh.getRange(startRow + 3, 3, n, n);
  data.setValues(matrix).setNumberFormat(format);
  const backgrounds = matrix.map(row => row.map(v => Number(v) < Number(base) ? FS05_CFG.COLOR_LOW : FS05_CFG.COLOR_WHITE));
  data.setBackgrounds(backgrounds);
  sh.getRange(startRow + 2, 2, n + 1, n + 1).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(startRow + 3, 1, n, n + 2).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(centerRow, centerCol).setBackground(FS05_CFG.COLOR_CENTER).setFontWeight('bold').setBorder(true, true, true, true, true, true, '#FF0000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function FS05_format_(sh) {
  sh.getRange(1, 1, 99, 13).setFontFamily('Times New Roman').setFontSize(10).setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.getRange('A1:M1').setFontSize(14).setFontWeight('bold').setFontColor('#FFFFFF').setBackground(FS05_CFG.COLOR_TITLE);
  sh.getRange('A2:M2').setFontStyle('italic').setBackground('#D9EAF7');
  sh.setColumnWidth(1, 90); sh.setColumnWidth(2, 90);
  for (let c = 3; c <= 13; c++) sh.setColumnWidth(c, 82);
  sh.setFrozenRows(2);
}

function FS05_indexRows_(table) {
  const out = {};
  table.values.forEach(row => {
    const key = String(row[table.index.masp] || '').trim().toUpperCase() + '|' + FS05_num_(row[table.index.thangso]);
    out[key] = row;
  });
  return out;
}

function FS05_readTable_(sheet) {
  const cols = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, cols).getDisplayValues()[0];
  const values = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, cols).getValues() : [];
  const index = {};
  headers.forEach((h, i) => { index[FS05_key_(h)] = i; });
  return { values, index };
}

function FS05_findBlock_(sheet, name) {
  const data = sheet.getDataRange().getDisplayValues();
  const target = FS05_key_(name);
  let blockRow = 0;
  for (let r = 0; r < data.length && !blockRow; r++) if (data[r].some(v => FS05_key_(v) === target)) blockRow = r + 1;
  if (!blockRow) throw new Error('Không tìm thấy block ' + name + ' tại 01A. Kỹ thuật.');
  const startRow = blockRow + 2;
  let endRow = startRow - 1;
  for (let r = startRow; r <= sheet.getLastRow(); r++) {
    const first = String(sheet.getRange(r, 1).getDisplayValue() || '').trim();
    if (!first) break;
    if (['THONG_TIN_CHUNG', 'CHI_PHI_CHUNG', 'SAN_PHAM', 'KE_HOACH_BAN_THU_TIEN', 'TIEN_DO_CHI_PHI'].includes(first)) break;
    endRow = r;
  }
  return { startRow, rowCount: Math.max(0, endRow - startRow + 1) };
}

function FS05_findSummaryRow_(sheet, label) {
  const target = FS05_key_(label);
  const values = sheet.getRange(1, 2, sheet.getLastRow(), 1).getDisplayValues();
  for (let i = 0; i < values.length; i++) if (FS05_key_(values[i][0]) === target) return i + 1;
  return 0;
}

function FS05_readInfoValue_(sheet, label) {
  const target = FS05_key_(label);
  const values = sheet.getDataRange().getValues();
  for (let i = 0; i < values.length; i++) if (FS05_key_(values[i][0]) === target) return values[i][1];
  return '';
}

function FS05_emptyMatrix_() {
  return FS05_CFG.FACTORS.map(() => FS05_CFG.FACTORS.map(() => 0));
}

function FS05_require_(index, required, sheetName) {
  const missing = required.filter(k => index[k] == null);
  if (missing.length) throw new Error('Sheet "' + sheetName + '" thiếu cột: ' + missing.join(', '));
}

function FS05_group_(value) {
  const k = FS05_key_(value);
  return k.includes('chothue') || k.includes('thue') ? 'Cho thuê' : 'Bán';
}

function FS05_num_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  const text = String(value == null ? '' : value).trim().replace(/\s/g, '');
  if (!text) return 0;
  const n = Number(text.includes(',') && text.includes('.') ? text.replace(/\./g, '').replace(',', '.') : text.replace(/,/g, ''));
  return isFinite(n) ? n : 0;
}

function FS05_rate_(value) {
  if (typeof value === 'number') return value > 1 ? value / 100 : value;
  const text = String(value == null ? '' : value).trim();
  if (!text) return 0;
  const n = FS05_num_(text.replace('%', ''));
  return text.includes('%') || n > 1 ? n / 100 : n;
}

function FS05_norm_(value) {
  return String(value == null ? '' : value).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/\s+/g, ' ').trim();
}

function FS05_key_(value) {
  return FS05_norm_(value).replace(/²/g, '2').replace(/\^2/g, '2').replace(/m\s*2/g, 'm2').replace(/[^a-z0-9]/g, '');
}
