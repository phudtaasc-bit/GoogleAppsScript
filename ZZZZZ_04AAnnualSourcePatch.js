/*************************************************
 * ZZZZZ_04AAnnualSourcePatch.gs
 * Chuẩn hóa Sheet 04A: toàn bộ số liệu tổng/năm lấy từ Sheet 04.
 * Các cấu phần chi phí chi tiết được phản chiếu vào cuối Sheet 04
 * từ Sheet 03 theo tháng, để Sheet 04 là nguồn tổng hợp duy nhất.
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

  const startCol = 43; // AQ
  if (sh04.getMaxColumns() < startCol + mirror.length - 1) {
    sh04.insertColumnsAfter(sh04.getMaxColumns(), startCol + mirror.length - 1 - sh04.getMaxColumns());
  }

  sh04.getRange(1, startCol, 1, mirror.length)
    .setValues([mirror.map(() => 'CHI PHÍ CHI TIẾT')]);
  sh04.getRange(2, startCol, 1, mirror.length)
    .setValues([mirror.map(x => x[0])]);

  const src03 = `'03. Chi phí & Vốn'`;
  mirror.forEach((x, i) => {
    const col = startCol + i;
    sh04.getRange(3, col).setFormula(
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
    sh04.getRange(3, 3, Math.max(0, lastRow04 - 2), 1).getValues()
      .flat()
      .filter(v => v !== '' && !isNaN(Number(v)))
      .map(Number)
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
    netAfterFinancing: colByHeader('Dòng tiền thuần sau tài trợ'),
    fcff: colByHeader('FCFF_TIPV'),
    fcfe: colByHeader('FCFE khả dụng cho CSH') || colByHeader('FCFE_EPV')
  };

  const required = [
    'customer','equity','loan','xdTb','gpmb','land','htkt','selling','contingency',
    'vatIn','vatPay','cit','interest','principal','netAfterFinancing','fcff','fcfe'
  ];
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
    ['C', 'DÒNG TIỀN HIỆU QUẢ'],
    ['1', 'FCFF dự án'],
    ['2', 'Dòng tiền thuần sau tài trợ'],
    ['3', '(-) Vốn CSH góp mới'],
    ['4', '(+) Lãi vay'],
    ['5', 'FCFE vốn CSH']
  ].map(r => [r[0], r[1], '', ...years.map(() => '')]);

  sh.getRange(4, 1, rows.length, lastCol).setValues(rows);

  const R = {
    customer: 5, equity: 6, loan: 7, inflowTotal: 8,
    xdTb: 10, gpmb: 11, land: 12, htkt: 13, selling: 14, contingency: 15,
    operating: 16, maintenance: 17, vatIn: 18, vatPay: 19, cit: 20,
    interest: 21, principal: 22, outflowTotal: 23,
    fcff: 25, netAfterFinancing: 26, equityDeduction: 27, interestAddBack: 28, fcfe: 29
  };

  const setSum = (row, srcCol) => {
    if (!srcCol) {
      sh.getRange(row, totalCol, 1, years.length + 1).setValue(0);
      return;
    }
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
      const c = yearStartCol + i;
      const L = FS04A_colLetter_(c);
      sh.getRange(row, c).setFormula(`=SUM(${L}${fromRow}:${L}${toRow})`);
    });
  };

  setSum(R.customer, source.customer);
  setSum(R.equity, source.equity);
  setSum(R.loan, source.loan);
  setRowTotal(R.inflowTotal, R.customer, R.loan);

  setSum(R.xdTb, source.xdTb);
  setSum(R.gpmb, source.gpmb);
  setSum(R.land, source.land);
  setSum(R.htkt, source.htkt);
  setSum(R.selling, source.selling);
  setSum(R.contingency, source.contingency);
  setSum(R.operating, source.operating);
  setSum(R.maintenance, source.maintenance);
  setSum(R.vatIn, source.vatIn);
  setSum(R.vatPay, source.vatPay);
  setSum(R.cit, source.cit);
  setSum(R.interest, source.interest);
  setSum(R.principal, source.principal);
  setRowTotal(R.outflowTotal, R.xdTb, R.principal);

  setSum(R.fcff, source.fcff);
  setSum(R.netAfterFinancing, source.netAfterFinancing);
  setSum(R.equityDeduction, source.equity);
  setSum(R.interestAddBack, source.interest);

  // Quan hệ trình bày bắt buộc:
  // FCFE = Dòng tiền thuần sau tài trợ - Vốn CSH góp mới + Lãi vay.
  sh.getRange(R.fcfe, totalCol).setFormula(
    `=C${R.netAfterFinancing}-C${R.equityDeduction}+C${R.interestAddBack}`
  );
  years.forEach((_, i) => {
    const c = yearStartCol + i;
    const L = FS04A_colLetter_(c);
    sh.getRange(R.fcfe, c).setFormula(
      `=${L}${R.netAfterFinancing}-${L}${R.equityDeduction}+${L}${R.interestAddBack}`
    );
  });

  SpreadsheetApp.flush();
  FS04A_formatAnnualSource_(sh, lastCol);
}

function FS04A_formatAnnualSource_(sh, lastCol) {
  const lastRow = 29;
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
  [8, 23, 25, 29].forEach(row => sh.getRange(row, 1, 1, lastCol).setFontWeight('bold'));

  sh.getRange(4, 3, lastRow - 3, lastCol - 2).setNumberFormat('#,##0.0').setHorizontalAlignment('right');
  sh.getRange(1, 1, lastRow, 2).setHorizontalAlignment('left');
  sh.getRange(3, 1, lastRow - 2, 1).setHorizontalAlignment('center');
  sh.setColumnWidth(1, 55);
  sh.setColumnWidth(2, 330);
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
