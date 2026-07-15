function FS_lapSheet04() { return FS_hoiTuTaiTro(); }

function FS_hoiTuTaiTro() {
  const ss = SpreadsheetApp.getActive();
  const tech = FS_getSheet_(ss, FS_CFG.SHEETS.TECH, FS_CFG.SHEETS.TECH_LEGACY);
  const cost = ss.getSheetByName(FS_CFG.SHEETS.COST);
  if (!tech || !cost) throw new Error('Cần lập 01A. Kỹ thuật và 03. Chi phí & Vốn.');

  const info = FS_readInfo_(tech);
  const months = FS_num_(info['Số tháng mô hình']);
  const loanRatio = FS_rate_(info['Tỷ lệ vốn vay']);
  const monthlyRate = Math.pow(1 + FS_rate_(info['Lãi suất vay năm']), 1 / 12) - 1;

  let interest = Array(months).fill(0);
  let result = null;
  let converged = false;

  for (let iter = 1; iter <= FS_CFG.CONVERGENCE.maxIterations; iter++) {
    FS_lapSheet03A_voiLaiVay_(interest);
    SpreadsheetApp.flush();
    result = FS04_build_(months, loanRatio, monthlyRate);
    const next = result.map(r => FS_num_(r[16]));
    const diff = Math.max.apply(null, next.map((v, i) => Math.abs(v - interest[i])));
    interest = next;

    if (diff <= FS_CFG.CONVERGENCE.tolerance) {
      converged = true;
      PropertiesService.getDocumentProperties().setProperty('FS_CONVERGENCE', JSON.stringify({
        converged: true, iterations: iter, maxDiff: diff, time: new Date().toISOString()
      }));
      break;
    }
  }

  if (!converged) throw new Error('Không hội tụ sau ' + FS_CFG.CONVERGENCE.maxIterations + ' vòng.');
  FS_lapSheet03A_voiLaiVay_(interest);
  FS04_write_(result);
  return result;
}

function FS04_build_(months, loanRatio, monthlyRate) {
  const ss = SpreadsheetApp.getActive();
  const cost = ss.getSheetByName(FS_CFG.SHEETS.COST);
  const profit = ss.getSheetByName(FS_CFG.SHEETS.PROFIT);
  const cRows = cost.getLastRow() > 1 ? cost.getRange(2, 1, cost.getLastRow() - 1, 20).getValues() : [];
  const pRows = profit.getLastRow() > 1 ? profit.getRange(2, 1, profit.getLastRow() - 1, 19).getValues() : [];

  const cBy = Array.from({ length: months }, () => ({ cash: 0, vatOut: 0, costBefore: 0, vatIn: 0, costAfter: 0 }));
  cRows.forEach(r => {
    const i = FS_num_(r[0]) - 1;
    if (i < 0 || i >= months) return;
    cBy[i].cash += FS_num_(r[7]);
    cBy[i].vatOut += FS_num_(r[8]);
    cBy[i].costBefore += FS_num_(r[17]);
    cBy[i].vatIn += FS_num_(r[18]);
    cBy[i].costAfter += FS_num_(r[19]);
  });

  const taxBy = Array(months).fill(0);
  const patBy = Array(months).fill(0);
  pRows.forEach(r => {
    const i = FS_num_(r[0]) - 1;
    if (i < 0 || i >= months) return;
    taxBy[i] += FS_num_(r[17]);
    patBy[i] += FS_num_(r[18]);
  });

  let cash = 0;
  let debt = 0;
  let vatCredit = 0;
  const out = [];

  for (let i = 0; i < months; i++) {
    const x = cBy[i];
    const vatPay = Math.max(0, x.vatOut - vatCredit - x.vatIn);
    const vatCreditEnd = Math.max(0, vatCredit + x.vatIn - x.vatOut);
    const tax = taxBy[i];
    const operating = x.cash - x.costAfter - vatPay - tax;
    const interest = debt * monthlyRate;
    const need = Math.max(0, -(cash + operating - interest));
    const equity = need * (1 - loanRatio);
    const draw = need * loanRatio;
    const available = cash + operating - interest + equity + draw;
    const repay = Math.min(debt + draw + interest, Math.max(0, available));
    const debtEnd = Math.max(0, debt + draw + interest - repay);
    const cashEnd = Math.max(0, available - repay);
    const fcff = operating;
    const fcfe = fcff + draw - repay;
    const date = cRows.find(r => FS_num_(r[0]) === i + 1)?.[1] || '';
    const year = date instanceof Date ? date.getFullYear() : '';
    const quarter = date instanceof Date ? 'Q' + Math.ceil((date.getMonth() + 1) / 3) + '/' + year : '';

    out.push([
      i + 1, date, year, quarter, x.cash, x.vatOut, x.costBefore, x.vatIn, x.costAfter,
      vatCredit, vatPay, vatCreditEnd, tax, patBy[i], operating, need,
      interest, equity, draw, repay, debtEnd, cashEnd, fcff, fcfe
    ]);
    cash = cashEnd;
    debt = debtEnd;
    vatCredit = vatCreditEnd;
  }
  return out;
}

function FS04_write_(rows) {
  const ss = SpreadsheetApp.getActive();
  const sh = FS_getOrCreateSheet_(ss, FS_CFG.SHEETS.CASH, FS_CFG.SHEETS.CASH_LEGACY);
  FS_resetSheet_(sh, rows.length + 1, 24);
  sh.getRange(1, 1, 1, 24).setValues([[
    'Tháng số', 'Tháng', 'Năm', 'Quý', 'Dòng tiền thu khách hàng', 'VAT đầu ra',
    'Tổng chi trước VAT', 'VAT đầu vào', 'Tổng chi sau VAT', 'VAT khấu trừ đầu kỳ',
    'VAT phải nộp', 'VAT khấu trừ cuối kỳ', 'Thuế TNDN', 'LNST',
    'Dòng tiền trước tài trợ', 'Nhu cầu vốn', 'Lãi vay', 'Vốn góp CSH',
    'Giải ngân vay', 'Trả gốc', 'Dư nợ cuối kỳ', 'Tiền cuối kỳ', 'FCFF', 'FCFE'
  ]]);
  if (rows.length) sh.getRange(2, 1, rows.length, 24).setValues(rows);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(4);
  sh.getRange(1, 1, 1, 24).setFontWeight('bold').setBackground('#ddebf7').setWrap(true);
  if (rows.length) {
    sh.getRange(2, 2, rows.length, 1).setNumberFormat('MM/yyyy');
    sh.getRange(2, 5, rows.length, 20).setNumberFormat('#,##0.00');
  }
  sh.autoResizeColumns(1, 24);
}
