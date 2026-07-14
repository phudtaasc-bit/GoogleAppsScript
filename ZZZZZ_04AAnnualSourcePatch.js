/*************************************************
 * ZZZZZ_04AAnnualSourcePatch.gs
 * Chuẩn hóa Sheet 04A: toàn bộ số liệu tổng/năm lấy từ Sheet 04.
 * Đồng thời trình bày trực tiếp quan hệ giữa dòng tiền vào, dòng tiền ra,
 * dòng tiền thuần sau tài trợ, FCFF và FCFE.
 *************************************************/

const FS04A_BASE_LEGACY_ = FS_lapSheet04A;

FS_lapSheet04A = function() {
  const ss = SpreadsheetApp.getActive();
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  if (!sh04) throw new Error('Không tìm thấy sheet "04. Dòng tiền & Lợi nhuận".');
  if (!sh03) throw new Error('Không tìm thấy sheet "03. Chi phí & Vốn".');

  FS04A_mirrorCostDetailToSheet04_(sh04, sh03);
  FS04A_buildAnnualFromSheet04_(ss, sh04);
};

function FS04A_mirrorCostDetailToSheet04_(sh04, sh03) {
  const lastRow = sh04.getLastRow();
  if (lastRow < 3) return;

  const mirror = [
    ['Chi XD/TB/khác trước VAT', 'H'],
    ['Chi GPMB trước VAT', 'I'],
    ['Tiền SDĐ/thuê đất trước VAT', 'J'],
    ['Chi HTKT trước VAT', 'K'],
    ['Chi phí bán hàng trước VAT', 'L'],
    ['Chi phí dự phòng trước VAT', 'M']
  ];

  const startCol = 43;
  if (sh04.getMaxColumns() < startCol + mirror.length - 1) {
    sh04.insertColumnsAfter(sh04.getMaxColumns(), startCol + mirror.length - 1 - sh04.getMaxColumns());
  }

  sh04.getRange(1, startCol, 1, mirror.length).setValues([mirror.map(() => 'CHI PHÍ CHI TIẾT')]);
  sh04.getRange(2, startCol, 1, mirror.length).setValues([mirror.map(x => x[0])]);

  const src03 = `'03. Chi phí & Vốn'`;
  mirror.forEach((x, i) => {
    sh04.getRange(3, startCol + i).setFormula(
      `=MAP(A3:A${lastRow};LAMBDA(t;IF(t="";"";SUMIF(${src03}!A:A;t;${src03}!${x[1]}:${x[1]}))))`
    );
  });
  sh04.getRange(3, startCol, lastRow - 2, mirror.length).setNumberFormat('#,##0');
}

function FS04A_buildAnnualFromSheet04_(ss, sh04) {
  let sh = ss.getSheetByName('04A. TH dòng tiền');
  if (!sh) sh = ss.insertSheet('04A. TH dòng tiền');

  sh.showRows(1, sh.getMaxRows());
  sh.showColumns(1, sh.getMaxColumns());
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.setFrozenRows(0);
  sh.setFrozenColumns(0);
  sh.clearContents();
  sh.clearFormats();
  sh.setHiddenGridlines(true);

  const lastRow04 = sh04.getLastRow();
  const years = [...new Set(
    sh04.getRange(3, 3, Math.max(0, lastRow04 - 2), 1).getValues().flat()
      .filter(v => v !== '' && !isNaN(Number(v))).map(Number)
  )].sort((a, b) => a - b);
  if (!years.length) throw new Error('Sheet 04 chưa có dữ liệu năm.');

  const headers04 = sh04.getRange(2, 1, 1, sh04.getLastColumn()).getDisplayValues()[0];
  const colByHeader = name => {
    const key = FS04A_headerKey_(name);
    const idx = headers04.findIndex(h => FS04A_headerKey_(h) === key);
    return idx >= 0 ? FS04A_colLetter_(idx + 1) : '';
  };

  const source = {
    customer: colByHeader('Dòng tiền huy động từ KH'),
    equity: colByHeader('Tổng dòng CSH vào dự án') || colByHeader('CSH góp mới'),
    loan: colByHeader('Giải ngân vay'),
    xdTb: colByHeader('Chi XD/TB/khác trước VAT'),
    gpmb: colByHeader('Chi GPMB trước VAT'),
    land: colByHeader('Tiền SDĐ/thuê đất trước VAT'),
    htkt: colByHeader('Chi HTKT trước VAT'),
    selling: colByHeader('Chi phí bán hàng trước VAT'),
    contingency: colByHeader('Chi phí dự phòng trước VAT'),
    operating: colByHeader('Chi phí vận hành thuê trước VAT'),
    maintenance: colByHeader('Chi phí bảo trì trước VAT'),
    vatIn: colByHeader('VAT đầu vào'),
    vatPay: colByHeader('VAT phải nộp'),
    cit: colByHeader('Thuế TNDN'),
    interest: colByHeader('Lãi vay vốn hóa'),
    principal: colByHeader('Trả gốc'),
    fcff: colByHeader('FCFF_TIPV'),
    fcfe: colByHeader('FCFE khả dụng cho CSH') || colByHeader('FCFE_EPV')
  };

  const required = ['customer','equity','loan','xdTb','gpmb','land','htkt','selling','contingency',
    'vatIn','vatPay','cit','interest','principal','fcff','fcfe'];
  const missing = required.filter(k => !source[k]);
  if (missing.length) throw new Error('Thiếu cột nguồn trên Sheet 04: ' + missing.join(', '));

  const yearStartCol = 4;
  const totalCol = 3;
  const lastCol = yearStartCol + years.length - 1;
  const src04 = `'04. Dòng tiền & Lợi nhuận'`;

  sh.getRange(1, 1).setValue('BẢNG TỔNG HỢP DÒNG TIỀN DỰ ÁN THEO NĂM');
  sh.getRange(3, 1, 1, lastCol).setValues([['TT', 'Nội dung', 'Tổng', ...years]]);

  const rows = [
    ['A', 'DÒNG TIỀN VÀO'],
    ['1', 'Dòng tiền huy động từ khách hàng'],
    ['2', 'Dòng tiền vốn CSH'],
    ['3', 'Dòng tiền vốn vay'],
    ['A', 'TỔNG DÒNG TIỀN VÀO'],
    ['B', 'DÒNG TIỀN RA'],
    ['1', 'Chi XD/TB/khác trước VAT'],
    ['2', 'Chi GPMB trước VAT'],
    ['3', 'Tiền SDĐ/thuê đất trước VAT'],
    ['4', 'Chi HTKT trước VAT'],
    ['5', 'Chi phí bán hàng trước VAT'],
    ['6', 'Chi phí dự phòng trước VAT'],
    ['7', 'Chi phí vận hành trước VAT'],
    ['8', 'Chi phí bảo trì trước VAT'],
    ['9', 'VAT đầu vào'],
    ['10', 'VAT phải nộp'],
    ['11', 'Thuế TNDN'],
    ['12', 'Lãi vay'],
    ['13', 'Trả gốc vay'],
    ['B', 'TỔNG DÒNG TIỀN RA'],
    ['C', 'DÒNG TIỀN HIỆU QUẢ VÀ ĐỐI CHIẾU'],
    ['1', 'Dòng tiền thuần sau tài trợ = Tổng dòng tiền vào - Tổng dòng tiền ra'],
    ['2', '(-) Vốn CSH góp mới'],
    ['3', '(-) Vốn vay giải ngân'],
    ['4', '(+) Trả gốc vay'],
    ['5', '(+) Lãi vay'],
    ['6', 'FCFF dự án'],
    ['7', 'FCFE vốn CSH']
  ].map(r => [r[0], r[1], '', ...years.map(() => '')]);
  sh.getRange(4, 1, rows.length, lastCol).setValues(rows);

  const R = {
    customer: 5, equity: 6, loan: 7, inflowTotal: 8,
    xdTb: 10, gpmb: 11, land: 12, htkt: 13, selling: 14, contingency: 15,
    operating: 16, maintenance: 17, vatIn: 18, vatPay: 19, cit: 20,
    interest: 21, principal: 22, outflowTotal: 23,
    netAfterFinancing: 25, equityDeduction: 26, loanDeduction: 27,
    principalAddBack: 28, interestAddBack: 29, fcff: 30, fcfe: 31
  };

  const setSum = (row, srcCol) => {
    sh.getRange(row, totalCol).setFormula(`=SUM(${src04}!${srcCol}3:${srcCol}${lastRow04})/1000000000`);
    years.forEach((y, i) => {
      sh.getRange(row, yearStartCol + i).setFormula(
        `=SUMIF(${src04}!C3:C${lastRow04};${y};${src04}!${srcCol}3:${srcCol}${lastRow04})/1000000000`
      );
    });
  };
  const setRowTotal = (row, fromRow, toRow) => {
    sh.getRange(row, totalCol).setFormula(`=SUM(C${fromRow}:C${toRow})`);
    years.forEach((_, i) => {
      const L = FS04A_colLetter_(yearStartCol + i);
      sh.getRange(row, yearStartCol + i).setFormula(`=SUM(${L}${fromRow}:${L}${toRow})`);
    });
  };
  const setFormulaAcross = (row, builder) => {
    sh.getRange(row, totalCol).setFormula(builder('C'));
    years.forEach((_, i) => {
      const L = FS04A_colLetter_(yearStartCol + i);
      sh.getRange(row, yearStartCol + i).setFormula(builder(L));
    });
  };

  setSum(R.customer, source.customer);
  setSum(R.equity, source.equity);
  setSum(R.loan, source.loan);
  setRowTotal(R.inflowTotal, R.customer, R.loan);

  setSum(R.xdTb, source.xdTb); setSum(R.gpmb, source.gpmb); setSum(R.land, source.land);
  setSum(R.htkt, source.htkt); setSum(R.selling, source.selling); setSum(R.contingency, source.contingency);
  setSum(R.operating, source.operating); setSum(R.maintenance, source.maintenance); setSum(R.vatIn, source.vatIn);
  setSum(R.vatPay, source.vatPay); setSum(R.cit, source.cit); setSum(R.interest, source.interest);
  setSum(R.principal, source.principal);
  setRowTotal(R.outflowTotal, R.xdTb, R.principal);

  setFormulaAcross(R.netAfterFinancing, L => `=${L}${R.inflowTotal}-${L}${R.outflowTotal}`);
  setSum(R.equityDeduction, source.equity);
  setSum(R.loanDeduction, source.loan);
  setSum(R.principalAddBack, source.principal);
  setSum(R.interestAddBack, source.interest);

  setFormulaAcross(R.fcff, L =>
    `=${L}${R.netAfterFinancing}-${L}${R.equityDeduction}-${L}${R.loanDeduction}+${L}${R.principalAddBack}+${L}${R.interestAddBack}`
  );
  setFormulaAcross(R.fcfe, L =>
    `=${L}${R.netAfterFinancing}-${L}${R.equityDeduction}+${L}${R.interestAddBack}`
  );

  SpreadsheetApp.flush();
  FS04A_formatAnnualSource_(sh, lastCol);
}

function FS04A_formatAnnualSource_(sh, lastCol) {
  const lastRow = 31;
  sh.setFrozenRows(3);
  sh.setFrozenColumns(0);
  sh.getRange(1, 1, lastRow, lastCol)
    .setFontFamily('Times New Roman').setFontSize(11)
    .setVerticalAlignment('middle').setWrap(true)
    .setBorder(true, true, true, true, true, true);

  sh.getRange(1, 1, 1, lastCol).clearContent().setBackground('#FFFFFF').setBorder(false, false, false, false, false, false);
  sh.getRange('A1').setValue('BẢNG TỔNG HỢP DÒNG TIỀN DỰ ÁN THEO NĂM')
    .setFontWeight('bold').setFontSize(14).setHorizontalAlignment('left').setWrap(false);
  sh.getRange(3, 1, 1, lastCol).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#A6A6A6').setFontColor('#FFFFFF');

  [4, 9, 24].forEach(row => sh.getRange(row, 1, 1, lastCol).setFontWeight('bold').setBackground('#FFC000').setFontColor('#000000'));
  [8, 23, 25, 30, 31].forEach(row => sh.getRange(row, 1, 1, lastCol).setFontWeight('bold'));

  sh.getRange(4, 3, lastRow - 3, lastCol - 2).setNumberFormat('#,##0.0').setHorizontalAlignment('right');
  sh.getRange(1, 1, lastRow, 2).setHorizontalAlignment('left');
  sh.getRange(3, 1, lastRow - 2, 1).setHorizontalAlignment('center');
  sh.setColumnWidth(1, 55);
  sh.setColumnWidth(2, 430);
  for (let c = 3; c <= lastCol; c++) sh.setColumnWidth(c, 115);
  for (let r = 1; r <= lastRow; r++) sh.setRowHeight(r, 24);
  sh.setRowHeight(1, 28);
  sh.setRowHeight(2, 8);
  sh.setRowHeight(3, 32);
}

function FS04A_headerKey_(v) {
  return String(v || '').toLowerCase().normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd')
    .replace(/²/g, '2').replace(/[^a-z0-9]/g, '');
}

function FS04A_colLetter_(col) {
  let out = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    col = Math.floor((col - 1) / 26);
  }
  return out;
}
