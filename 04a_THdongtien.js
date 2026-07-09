/*************************************************
 * 04A_TH_DongTien.gs
 * Tổng hợp dòng tiền theo năm
 * Bỏ mục D - Số dư cuối kỳ
 *************************************************/

function FS_lapSheet04A() {
  const ss = SpreadsheetApp.getActive();

  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');

  if (!sh04) throw new Error('Không tìm thấy sheet "04. Dòng tiền & Lợi nhuận".');
  if (!sh03) throw new Error('Không tìm thấy sheet "03. Chi phí & Vốn".');

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

  const years = [...new Set(
    sh04.getRange('C3:C').getValues()
      .flat()
      .filter(v => v !== '' && !isNaN(Number(v)))
      .map(v => Number(v))
  )].sort();

  if (years.length === 0) throw new Error('Sheet 04 chưa có dữ liệu năm.');

  const yearStartCol = 4;
  const totalCol = 3;
  const lastCol = yearStartCol + years.length - 1;

  sh.getRange(1, 1).setValue('BẢNG TỔNG HỢP DÒNG TIỀN DỰ ÁN THEO NĂM');

  sh.getRange(3, 1, 1, lastCol).setValues([[
    'TT', 'Nội dung', 'Tổng', ...years
  ]]);

  const rows = [
    ['A', 'DÒNG TIỀN VÀO', '', ...years.map(() => '')],
    ['1', 'Dòng tiền huy động từ khách hàng', '', ...years.map(() => '')],
    ['2', 'Dòng tiền vốn CSH', '', ...years.map(() => '')],
    ['3', 'Dòng tiền vốn vay', '', ...years.map(() => '')],
    ['A', 'TỔNG DÒNG TIỀN VÀO', '', ...years.map(() => '')],

    ['B', 'DÒNG TIỀN RA', '', ...years.map(() => '')],
    ['1', 'Chi XD/TB/khác trước VAT', '', ...years.map(() => '')],
    ['2', 'Chi GPMB trước VAT', '', ...years.map(() => '')],
    ['3', 'Tiền SDĐ/thuê đất trước VAT', '', ...years.map(() => '')],
    ['4', 'Chi HTKT trước VAT', '', ...years.map(() => '')],
    ['5', 'Chi phí bán hàng trước VAT', '', ...years.map(() => '')],
    ['6', 'Chi phí dự phòng trước VAT', '', ...years.map(() => '')],
    ['7', 'VAT đầu vào', '', ...years.map(() => '')],
    ['8', 'VAT phải nộp', '', ...years.map(() => '')],
    ['9', 'Thuế TNDN', '', ...years.map(() => '')],
    ['10', 'Lãi vay', '', ...years.map(() => '')],
    ['11', 'Trả gốc vay', '', ...years.map(() => '')],
    ['B', 'TỔNG DÒNG TIỀN RA', '', ...years.map(() => '')],

    ['C', 'DÒNG TIỀN HIỆU QUẢ', '', ...years.map(() => '')],
    ['1', 'FCFF dự án', '', ...years.map(() => '')],
    ['2', 'FCFE vốn CSH', '', ...years.map(() => '')]
  ];

  sh.getRange(4, 1, rows.length, lastCol).setValues(rows);

  const src04 = `'04. Dòng tiền & Lợi nhuận'`;
  const src03 = `'03. Chi phí & Vốn'`;

  const r = {
    customer: 5,
    csh: 6,
    loan: 7,
    inflowTotal: 8,

    xdTb: 10,
    gpmb: 11,
    tienSd: 12,
    htkt: 13,
    banHang: 14,
    duPhong: 15,
    vatIn: 16,
    vatPay: 17,
    cit: 18,
    interest: 19,
    principal: 20,
    outflowTotal: 21,

    fcff: 23,
    fcfe: 24
  };

  function setSumByYear(row, sheetRef, yearCol, srcCol) {
    sh.getRange(row, totalCol)
      .setFormula(`=SUM(${sheetRef}!${srcCol}3:${srcCol})/1000000000`);

    years.forEach((y, i) => {
      sh.getRange(row, yearStartCol + i)
        .setFormula(`=SUMIF(${sheetRef}!${yearCol}:${yearCol};${y};${sheetRef}!${srcCol}:${srcCol})/1000000000`);
    });
  }

  function setRowTotal(row, fromRow, toRow) {
    sh.getRange(row, totalCol).setFormula(`=SUM(C${fromRow}:C${toRow})`);
    years.forEach((_, i) => {
      const c = yearStartCol + i;
      const col = FS04A_colLetter_(c);
      sh.getRange(row, c).setFormula(`=SUM(${col}${fromRow}:${col}${toRow})`);
    });
  }

  setSumByYear(r.customer, src04, 'C', 'G');
  setSumByYear(r.csh, src04, 'C', 'S');
  setSumByYear(r.loan, src04, 'C', 'V');
  setRowTotal(r.inflowTotal, r.customer, r.loan);

  setSumByYear(r.xdTb, src03, 'C', 'H');
  setSumByYear(r.gpmb, src03, 'C', 'I');
  setSumByYear(r.tienSd, src03, 'C', 'J');
  setSumByYear(r.htkt, src03, 'C', 'K');
  setSumByYear(r.banHang, src03, 'C', 'L');
  setSumByYear(r.duPhong, src03, 'C', 'M');
  setSumByYear(r.vatIn, src03, 'C', 'O');

  setSumByYear(r.vatPay, src04, 'C', 'K');
  setSumByYear(r.cit, src04, 'C', 'L');
  setSumByYear(r.interest, src04, 'C', 'W');
  setSumByYear(r.principal, src04, 'C', 'X');
  setRowTotal(r.outflowTotal, r.xdTb, r.principal);

  setSumByYear(r.fcff, src04, 'C', 'AJ');
  setSumByYear(r.fcfe, src04, 'C', 'AK');

  SpreadsheetApp.flush();
  FS04A_format_(sh, lastCol);
}

function FS04A_format_(sh, lastCol) {
  const lastRow = 24;

  sh.setFrozenRows(3);
  sh.setFrozenColumns(0);

  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();

  sh.getRange(1, 1, lastRow, lastCol)
    .setFontFamily('Times New Roman')
    .setFontSize(11)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true);

  sh.getRange(1, 1, 1, lastCol)
    .clearContent()
    .setBackground('#FFFFFF')
    .setFontColor('#000000')
    .setBorder(false, false, false, false, false, false);

  sh.getRange('A1')
    .setValue('BẢNG TỔNG HỢP DÒNG TIỀN DỰ ÁN THEO NĂM')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle')
    .setWrap(false)
    .setBackground('#FFFFFF')
    .setFontColor('#000000');

  sh.getRange(3, 1, 1, lastCol)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBackground('#A6A6A6')
    .setFontColor('#FFFFFF')
    .setBorder(true, true, true, true, true, true);

  [4, 9, 22].forEach(row => {
    sh.getRange(row, 1, 1, lastCol)
      .setFontWeight('bold')
      .setBackground('#FFC000')
      .setFontColor('#000000');
  });

  [8, 21, 23, 24].forEach(row => {
    sh.getRange(row, 1, 1, lastCol).setFontWeight('bold');
  });

  sh.getRange(4, 3, lastRow - 3, lastCol - 2)
    .setNumberFormat('#,##0.0')
    .setHorizontalAlignment('right');

  sh.getRange(1, 1, lastRow, 2).setHorizontalAlignment('left');
  sh.getRange(3, 1, lastRow - 2, 1).setHorizontalAlignment('center');

  sh.setColumnWidth(1, 55);
  sh.setColumnWidth(2, 330);
  for (let c = 3; c <= lastCol; c++) sh.setColumnWidth(c, 115);

  for (let row = 1; row <= lastRow; row++) sh.setRowHeight(row, 24);

  sh.setRowHeight(1, 28);
  sh.setRowHeight(2, 8);
  sh.setRowHeight(3, 32);

  sh.getRange(1, 1, lastRow, lastCol).setWrap(true);
  sh.getRange('A1').setWrap(false);
}

function FS04A_colLetter_(col) {
  let temp = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    temp = String.fromCharCode(65 + rem) + temp;
    col = Math.floor((col - rem - 1) / 26);
  }
  return temp;
}