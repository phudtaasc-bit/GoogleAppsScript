/*************************************************
 * 05_DoNhay_Fast.gs
 * Độ nhạy nhanh - NPV/IRR dự án theo FCFF
 * Không rebuild Sheet 02/03/04 theo từng kịch bản
 *************************************************/

const FS05F = {
  SHEET: '05. Độ nhạy',
  REV_FACTORS: [-0.10, -0.08, -0.05, -0.04, -0.02, 0, 0.02, 0.04, 0.05, 0.08, 0.10],
  COST_FACTORS: [-0.10, -0.08, -0.05, -0.04, -0.02, 0, 0.02, 0.04, 0.05, 0.08, 0.10],
  COLOR_BASE: '#008000',
  COLOR_CENTER: '#FFFF00',
  COLOR_LOW: '#E6B8B7',
  COLOR_WHITE: '#FFFFFF',
  COLOR_TITLE: '#1F4E78',
  COLOR_HEADER: '#A6A6A6',
  COLOR_SECTION: '#FFC000'
};

function FS05_DoNhay_Fast() {
  const ss = SpreadsheetApp.getActive();
  const input = FS05F_readInput_();

  let sh = ss.getSheetByName(FS05F.SHEET);
  if (!sh) sh = ss.insertSheet(FS05F.SHEET);

  sh.clear();
  sh.clearFormats();

  const baseSheet = ss.getSheetByName('00. Tổng hợp');
  if (!baseSheet) throw new Error('Không tìm thấy sheet "00. Tổng hợp".');

  const base = {
    npvProject: Number(baseSheet.getRange('D33').getValue()) || 0,
    irrProject: Number(baseSheet.getRange('D34').getValue()) || 0
  };

  FS05F_layout_(sh);

  const matrix = FS05F_computeMatrices_(input, base);

  FS05F_writeBlock_(sh, 4, 'NPV', base.npvProject, matrix.npvProject, '#,##0.0');
  FS05F_writeBlock_(sh, 20, 'IRR', base.irrProject, matrix.irrProject, '0.0%');

  FS05F_format_(sh);

  SpreadsheetApp.getUi().alert('Đã chạy xong độ nhạy nhanh: NPV/IRR dự án.');
}

/***********************
 * READ INPUT
 ***********************/

function FS05F_readInput_() {
  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName('01. Kỹ thuật');
  if (!tech) throw new Error('Không tìm thấy sheet "01. Kỹ thuật".');

  const sp = FS05F_getBlockInfo_(tech, 'SAN_PHAM');
  const kh = FS05F_getBlockInfo_(tech, 'KE_HOACH_BAN_THU_TIEN');
  const cp = FS05F_getBlockInfo_(tech, 'CHI_PHI_CHUNG');
  const td = FS05F_getBlockInfo_(tech, 'TIEN_DO_CHI_PHI');

  const cfg = {
    months: Number(FS05F_getInfoValue_(tech, 'Số tháng mô hình')) || 40,
    discountRate: Number(FS05F_getInfoValue_(tech, 'Tỷ suất chiết khấu')) || 0,
    costInflation: Number(FS05F_getInfoValue_(tech, 'Tỷ lệ trượt chi phí/năm')) || 0
  };

  const cpRows = tech.getRange(cp.startRow, 1, cp.rowCount, 6).getValues().map(r => ({
    name: String(r[0] || ''),
    beforeVat: Number(r[1]) || 0,
    vatRate: FS05F_percent_(r[2]),
    afterVat: Number(r[3]) || 0,
    note: String(r[4] || ''),
    ratio: FS05F_percent_(r[5])
  }));

  const products = tech.getRange(sp.startRow, 1, sp.rowCount, 12).getValues().map(r => ({
    name: String(r[0] || ''),
    form: String(r[1] || ''),
    area: Number(r[2]) || 0,
    salePrice: Number(r[3]) || 0,
    rentPrice: Number(r[4]) || 0,
    cpxdM2: Number(r[5]) || 0,
    vatOut: FS05F_percent_(r[6]),
    taxRate: FS05F_percent_(r[7]),
    occupancy: FS05F_percent_(r[8]),
    cpvh: FS05F_percent_(r[9]),
    note: String(r[10] || ''),
    landArea: Number(r[11]) || 0
  }));

  const planRows = tech.getRange(kh.startRow, 1, kh.rowCount, 8).getValues().map(r => ({
    group: String(r[0] || ''),
    product: String(r[1] || ''),
    phase: Number(r[3]) || 0,
    start: Number(r[4]) || 0,
    duration: Number(r[5]) || 0,
    rate: FS05F_percent_(r[6]),
    note: String(r[7] || '')
  }));

  const scheduleRows = tech.getRange(td.startRow, 1, td.rowCount, 5).getValues().map(r => ({
    item: String(r[0] || ''),
    start: Number(r[1]) || 0,
    duration: Number(r[2]) || 0,
    rate: FS05F_percent_(r[3]),
    type: String(r[4] || '')
  }));

  return { cfg, cpRows, products, planRows, scheduleRows };
}

/***********************
 * MATRIX ENGINE
 ***********************/

function FS05F_computeMatrices_(input, base) {
  const out = {
    npvProject: [],
    irrProject: []
  };

  FS05F.COST_FACTORS.forEach(costChange => {
    const rowNPV = [];
    const rowIRR = [];

    FS05F.REV_FACTORS.forEach(revChange => {
      if (revChange === 0 && costChange === 0) {
        rowNPV.push(base.npvProject);
        rowIRR.push(base.irrProject);
      } else {
        const kpi = FS05F_computeScenario_(input, 1 + revChange, 1 + costChange).kpi;
        rowNPV.push(kpi.npvProject / 1000000000);
        rowIRR.push(kpi.irrProject);
      }
    });

    out.npvProject.push(rowNPV);
    out.irrProject.push(rowIRR);
  });

  return out;
}

function FS05F_computeScenario_(input, revFactor, costFactor) {
  const revenue = FS05F_computeRevenue_(input, revFactor);
  const cost = FS05F_computeCosts_(input, costFactor, revenue);
  const tax = FS05F_computeTax_(input, revenue, cost);
  const cash = FS05F_computeFCFF_(input, revenue, cost, tax);
  const kpi = FS05F_calcKpi_(input, cash);

  return { revenue, cost, tax, cash, kpi };
}

/***********************
 * REVENUE
 ***********************/

function FS05F_computeRevenue_(input, revFactor) {
  const out = [];

  for (let t = 1; t <= input.cfg.months; t++) {
    let beforeVat = 0;
    let vatOut = 0;
    let cashInVat = 0;

    input.products.forEach(p => {
      const form = FS05F_norm_(p.form);
      const progress = FS05F_collectProgress_(input.planRows, p.name, t);
      let amount = 0;

      if (form.includes('ban')) {
        amount = p.area * p.salePrice * revFactor * progress;
      } else if (form.includes('thue')) {
        amount = p.area * p.rentPrice * p.occupancy * revFactor;
      }

      beforeVat += amount;
      vatOut += amount * p.vatOut;
      cashInVat += amount * (1 + p.vatOut);
    });

    out.push({ beforeVat, vatOut, cashInVat });
  }

  return out;
}

function FS05F_collectProgress_(plans, productName, month) {
  const target = FS05F_norm_(productName);

  return plans
    .filter(p => FS05F_norm_(p.group) === 'thu tien' && FS05F_norm_(p.product) === target)
    .reduce((s, p) => {
      if (p.duration <= 0) return s;
      if (month >= p.start && month < p.start + p.duration) return s + p.rate / p.duration;
      return s;
    }, 0);
}

/***********************
 * COST + VAT
 ***********************/

function FS05F_computeCosts_(input, costFactor, revenue) {
  const rows = [];

  const salesRatio = FS05F_getCostRatio_(input, 'Chi phí bán hàng');
  const contingencyRatio = FS05F_getCostRatio_(input, 'Chi phí dự phòng');

  for (let t = 1; t <= input.cfg.months; t++) {
    const xd = FS05F_costBySchedule_(input, 'Chi phí XD/TB/khác', t, true, costFactor);
    const gpmb = FS05F_costBySchedule_(input, 'Chi phí GPMB', t, false, costFactor);
    const land = FS05F_costBySchedule_(input, 'Tiền SDĐ/thuê đất', t, false, costFactor);
    const htkt = FS05F_costBySchedule_(input, 'Chi phí HTKT', t, true, costFactor);

    const selling = revenue[t - 1].beforeVat * salesRatio;
    const contingency = (xd + htkt) * contingencyRatio;

    const vatIn =
      xd * FS05F_getVatRate_(input, 'Chi phí XD/TB/khác') +
      gpmb * FS05F_getVatRate_(input, 'Chi phí GPMB') +
      land * FS05F_getVatRate_(input, 'Tiền SDĐ/thuê đất') +
      htkt * FS05F_getVatRate_(input, 'Chi phí HTKT') +
      selling * FS05F_getVatRate_(input, 'Chi phí bán hàng') +
      contingency * FS05F_getVatRate_(input, 'Chi phí dự phòng');

    const beforeVat = xd + gpmb + land + htkt + selling + contingency;
    const afterVat = beforeVat + vatIn;

    rows.push({
      xd,
      gpmb,
      land,
      htkt,
      selling,
      contingency,
      vatIn,
      beforeVat,
      afterVat,
      vatPayable: 0
    });
  }

  FS05F_applyVat_(rows, revenue);

  return rows;
}

function FS05F_costBySchedule_(input, item, month, applyInflation, costFactor) {
  const total = FS05F_getCostBeforeVat_(input, item) * costFactor;
  const schedules = input.scheduleRows.filter(r => FS05F_norm_(r.item) === FS05F_norm_(item));

  let out = 0;

  schedules.forEach(s => {
    const duration = Math.max(1, s.duration);

    if (month >= s.start && month < s.start + duration) {
      const base = total * s.rate / duration;
      const infl = applyInflation ? Math.pow(1 + input.cfg.costInflation, (month - 1) / 12) : 1;
      out += base * infl;
    }
  });

  return out;
}

function FS05F_applyVat_(costRows, revenueRows) {
  let credit = 0;

  for (let i = 0; i < costRows.length; i++) {
    const vatOut = revenueRows[i].vatOut;
    const vatIn = costRows[i].vatIn;

    if (i === costRows.length - 1) {
      costRows[i].vatPayable = vatOut - credit - vatIn;
      credit = 0;
    } else {
      costRows[i].vatPayable = Math.max(0, vatOut - credit - vatIn);
      credit = Math.max(0, credit + vatIn - vatOut);
    }
  }
}

/***********************
 * TAX
 ***********************/

function FS05F_computeTax_(input, revenue, cost) {
  const tax = Array(input.cfg.months).fill(0);

  const totalCost = cost.reduce((s, r) =>
    s + r.xd + r.gpmb + r.land + r.htkt + r.selling + r.contingency, 0
  );

  const weights = revenue.map(r => Math.max(0, r.beforeVat));
  const totalWeight = weights.reduce((s, v) => s + v, 0) || 1;

  for (let i = 0; i < input.cfg.months; i++) {
    const costAlloc = totalCost * weights[i] / totalWeight;
    const rate = FS05F_avgTaxRateByMonth_(input, i + 1);
    const taxable = Math.max(0, revenue[i].beforeVat - costAlloc);
    tax[i] = taxable * rate;
  }

  return tax;
}

function FS05F_avgTaxRateByMonth_(input, month) {
  let rev = 0;
  let weightedTax = 0;

  input.products.forEach(p => {
    const progress = FS05F_collectProgress_(input.planRows, p.name, month);
    const form = FS05F_norm_(p.form);
    let q = 0;

    if (form.includes('ban')) {
      q = p.area * p.salePrice * progress;
    } else if (form.includes('thue')) {
      q = p.area * p.rentPrice * p.occupancy;
    }

    rev += q;
    weightedTax += q * p.taxRate;
  });

  return rev === 0 ? 0 : weightedTax / rev;
}

/***********************
 * FCFF / KPI
 ***********************/

function FS05F_computeFCFF_(input, revenue, cost, tax) {
  const fcff = [];

  for (let i = 0; i < input.cfg.months; i++) {
    const value =
      revenue[i].cashInVat -
      cost[i].afterVat -
      cost[i].vatPayable -
      tax[i];

    fcff.push(value);
  }

  return { fcff };
}

function FS05F_calcKpi_(input, cash) {
  const monthlyDiscount = Math.pow(1 + input.cfg.discountRate, 1 / 12) - 1;

  return {
    npvProject: FS05F_npv_(cash.fcff, monthlyDiscount),
    irrProject: FS05F_annualIrr_(cash.fcff)
  };
}

function FS05F_npv_(arr, rate) {
  return arr.reduce((s, v, i) => s + v / Math.pow(1 + rate, i + 1), 0);
}

function FS05F_annualIrr_(arr) {
  const r = FS05F_irr_(arr);
  return r === null ? 0 : Math.pow(1 + r, 12) - 1;
}

function FS05F_irr_(values) {
  if (!values.some(v => v > 0) || !values.some(v => v < 0)) return null;

  let rate = 0.02;

  for (let i = 0; i < 100; i++) {
    let f = 0;
    let df = 0;

    values.forEach((v, idx) => {
      const t = idx + 1;
      f += v / Math.pow(1 + rate, t);
      df += -t * v / Math.pow(1 + rate, t + 1);
    });

    if (Math.abs(df) < 1e-12) break;

    const next = rate - f / df;
    if (!isFinite(next)) break;
    if (Math.abs(next - rate) < 1e-8) return next;

    rate = next;
  }

  return rate;
}

/***********************
 * WRITE / FORMAT
 ***********************/

function FS05F_layout_(sh) {
  sh.getRange('A1:M1').merge().setValue('05. BẢNG PHÂN TÍCH ĐỘ NHẠY');
  sh.getRange('A2:M2').merge().setValue('Màu hồng: kết quả thấp hơn Base. Màu trắng: kết quả lớn hơn hoặc bằng Base.');
}

function FS05F_writeBlock_(sh, startRow, title, base, matrix, fmt) {
  const n = FS05F.REV_FACTORS.length;
  const m = FS05F.COST_FACTORS.length;
  const centerRow = startRow + 8;
  const centerCol = 8;

  sh.getRange(startRow, 2)
    .setValue(title)
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center');

  sh.getRange(startRow + 1, 3, 1, n)
    .merge()
    .setValue('Tăng/giảm giá bán')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sh.getRange(startRow + 2, 2)
    .setValue(base)
    .setNumberFormat(fmt)
    .setBackground(FS05F.COLOR_BASE)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('right');

  sh.getRange(startRow + 2, 3, 1, n)
    .setValues([FS05F.REV_FACTORS])
    .setNumberFormat('0%')
    .setBackground(FS05F.COLOR_HEADER)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sh.getRange(startRow + 3, 1, m, 1)
    .merge()
    .setValue('Tăng/giảm\nvốn đầu tư')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sh.getRange(startRow + 3, 2, m, 1)
    .setValues(FS05F.COST_FACTORS.map(x => [x]))
    .setNumberFormat('0%')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sh.getRange(startRow + 3, 3, m, n)
    .setValues(matrix)
    .setNumberFormat(fmt)
    .setHorizontalAlignment('right');

  sh.getRange(startRow + 2, 2, m + 1, n + 1)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange(startRow + 3, 1, m, n + 2)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  FS05F_colorByBase_(sh, startRow);

  sh.getRange(centerRow, centerCol)
    .setBackground(FS05F.COLOR_CENTER)
    .setFontWeight('bold')
    .setBorder(true, true, true, true, true, true, '#FF0000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function FS05F_colorByBase_(sh, startRow) {
  const base = Number(sh.getRange(startRow + 2, 2).getValue()) || 0;

  const range = sh.getRange(
    startRow + 3,
    3,
    FS05F.COST_FACTORS.length,
    FS05F.REV_FACTORS.length
  );

  const values = range.getValues();

  const backgrounds = values.map(row =>
    row.map(v => Number(v) < base ? FS05F.COLOR_LOW : FS05F.COLOR_WHITE)
  );

  range.setBackgrounds(backgrounds);
}

function FS05F_format_(sh) {
  const lastRow = 33;
  const lastCol = 13;

  sh.getRange(1, 1, lastRow, lastCol)
    .setFontFamily('Times New Roman')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');

  sh.getRange('A1:M1')
    .setFontSize(14)
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground(FS05F.COLOR_TITLE);

  sh.getRange('A2:M2')
    .setFontStyle('italic')
    .setBackground('#D9EAF7');

  [4, 20].forEach(r => {
    sh.getRange(r, 1, 1, lastCol)
      .setBackground(FS05F.COLOR_SECTION)
      .setFontWeight('bold');
  });

  sh.setColumnWidth(1, 90);
  sh.setColumnWidth(2, 90);
  for (let c = 3; c <= 13; c++) sh.setColumnWidth(c, 90);

  sh.setFrozenRows(2);
}

/***********************
 * HELPERS
 ***********************/

function FS05F_getBlockInfo_(sheet, blockName) {
  const blockRow = FS05F_findRow_(sheet, blockName);
  if (!blockRow) throw new Error(`Không tìm thấy block ${blockName}`);

  const startRow = blockRow + 2;
  let endRow = startRow;
  let blank = 0;

  for (let r = startRow; r <= sheet.getLastRow(); r++) {
    const firstCol = String(sheet.getRange(r, 1).getDisplayValue() || '').trim();

    if (['CHI_PHI_CHUNG', 'SAN_PHAM', 'KE_HOACH_BAN_THU_TIEN', 'TIEN_DO_CHI_PHI'].includes(firstCol)) break;

    if (!firstCol) {
      blank++;
      if (blank >= 3) {
        endRow = r - blank;
        break;
      }
    } else {
      blank = 0;
      endRow = r;
    }
  }

  return {
    blockRow,
    startRow,
    endRow,
    rowCount: Math.max(0, endRow - startRow + 1)
  };
}

function FS05F_findRow_(sheet, text) {
  const data = sheet.getDataRange().getDisplayValues();
  const target = FS05F_norm_(text);

  for (let r = 0; r < data.length; r++) {
    if (data[r].some(v => FS05F_norm_(v) === target)) return r + 1;
  }

  return null;
}

function FS05F_getInfoValue_(sheet, label) {
  const row = FS05F_findRow_(sheet, label);
  return row ? sheet.getRange(row, 2).getValue() : '';
}

function FS05F_getCostBeforeVat_(input, item) {
  const r = input.cpRows.find(x => FS05F_norm_(x.name) === FS05F_norm_(item));
  return r ? r.beforeVat : 0;
}

function FS05F_getVatRate_(input, item) {
  const r = input.cpRows.find(x => FS05F_norm_(x.name) === FS05F_norm_(item));
  return r ? r.vatRate : 0;
}

function FS05F_getCostRatio_(input, item) {
  const r = input.cpRows.find(x => FS05F_norm_(x.name) === FS05F_norm_(item));
  return r ? r.ratio : 0;
}

function FS05F_percent_(v) {
  if (typeof v === 'number') return v > 1 ? v / 100 : v;
  const s = String(v || '').replace('%', '').replace(',', '.').trim();
  const n = Number(s);
  if (!isFinite(n)) return 0;
  return n > 1 ? n / 100 : n;
}

function FS05F_norm_(v) {
  return String(v || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/\s+/g, ' ')
    .trim();
}

