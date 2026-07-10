/*************************************************
 * 03B_FinanceStateSync.gs
 * Đồng bộ toàn bộ trạng thái tài trợ từ Sheet 04 về Sheet 03.
 * Sheet 04 là nguồn tính duy nhất cho cash waterfall và dư nợ.
 *************************************************/

function FS03_capNhatNguonVonTuSheet04_V2() {
  const ss = SpreadsheetApp.getActive();
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');

  if (!sh03 || !sh04) {
    throw new Error('Không tìm thấy Sheet 03 hoặc Sheet 04 để đồng bộ trạng thái tài trợ.');
  }

  const rowCount03 = Math.max(0, sh03.getLastRow() - 1);
  const rowCount04 = Math.max(0, sh04.getLastRow() - 2);
  if (!rowCount03 || !rowCount04) {
    throw new Error('Sheet 03 hoặc Sheet 04 chưa có dữ liệu để đồng bộ trạng thái tài trợ.');
  }

  // Sheet 04: A:Y để lấy tháng và các cột Q:Y.
  const rows04 = sh04.getRange(3, 1, rowCount04, 25).getValues();
  const stateByMonth = {};

  rows04.forEach(row => {
    const month = FS03Sync_num_(row[0]);
    if (!month) return;

    stateByMonth[month] = {
      fundingNeed: FS03Sync_num_(row[16]),       // Q - Nhu cầu vốn
      equityNew: FS03Sync_num_(row[18]),         // S - CSH góp mới
      loanDraw: FS03Sync_num_(row[21]),          // V - Giải ngân vay
      interest: FS03Sync_num_(row[22]),          // W - Lãi vay vốn hóa
      principal: FS03Sync_num_(row[23]),         // X - Trả gốc
      debtEnd: FS03Sync_num_(row[24])            // Y - Dư nợ cuối kỳ
    };
  });

  const months03 = sh03.getRange(2, 1, rowCount03, 1).getValues();
  const out = [];
  let cumulativeEquity = 0;
  let cumulativeDraw = 0;
  let cumulativePrincipal = 0;
  let previousCash = 0;

  // Tiền cuối kỳ Sheet 04 nằm ở Z, đọc riêng để giữ đúng cash balance.
  const cash04 = sh04.getRange(3, 26, rowCount04, 1).getValues();
  const cashByMonth = {};
  rows04.forEach((row, index) => {
    const month = FS03Sync_num_(row[0]);
    if (month) cashByMonth[month] = FS03Sync_num_(cash04[index][0]);
  });

  months03.forEach(row => {
    const month = FS03Sync_num_(row[0]);
    const s = stateByMonth[month] || {
      fundingNeed: 0,
      equityNew: 0,
      loanDraw: 0,
      interest: 0,
      principal: 0,
      debtEnd: 0
    };

    cumulativeEquity += s.equityNew;
    cumulativeDraw += s.loanDraw;
    cumulativePrincipal += s.principal;

    const cashEnd = cashByMonth[month] || 0;
    const cashMovement = cashEnd - previousCash;

    // U:AE = 11 cột tài trợ trên Sheet 03.
    out.push([
      s.fundingNeed,          // U - Nhu cầu vốn
      s.equityNew,            // V - Vốn góp CSH
      cumulativeEquity,       // W - Lũy kế vốn góp CSH
      s.loanDraw,             // X - Giải ngân vay
      cumulativeDraw,         // Y - Lũy kế giải ngân vay
      s.interest,             // Z - Lãi vay vốn hóa
      s.principal,            // AA - Trả gốc
      cumulativePrincipal,    // AB - Lũy kế trả gốc
      s.debtEnd,              // AC - Dư nợ cuối kỳ
      cashMovement,           // AD - Biến động tiền sau tài trợ
      cashEnd                 // AE - Tiền cuối kỳ
    ]);

    previousCash = cashEnd;
  });

  sh03.getRange(2, 21, out.length, 11).setValues(out);
  SpreadsheetApp.flush();
  FS03Sync_assertWrittenState_(sh03, months03, out);
}

function FS03Sync_assertWrittenState_(sh03, months03, expected) {
  const actual = sh03.getRange(2, 21, expected.length, 11).getValues();
  const tolerance = 10;
  const labels = [
    'Nhu cầu vốn', 'Vốn góp CSH', 'Lũy kế vốn góp CSH',
    'Giải ngân vay', 'Lũy kế giải ngân vay', 'Lãi vay',
    'Trả gốc', 'Lũy kế trả gốc', 'Dư nợ cuối kỳ',
    'Dòng tiền sau tài trợ', 'Tiền cuối kỳ'
  ];
  const errors = [];

  for (let r = 0; r < expected.length; r++) {
    const month = FS03Sync_num_(months03[r][0]) || r + 1;
    for (let c = 0; c < 11; c++) {
      const delta = Math.abs(FS03Sync_num_(actual[r][c]) - FS03Sync_num_(expected[r][c]));
      if (delta > tolerance) {
        errors.push(`Tháng ${month}: ${labels[c]} lệch ${Math.round(delta).toLocaleString('vi-VN')} đồng.`);
        if (errors.length >= 20) break;
      }
    }
    if (errors.length >= 20) break;
  }

  if (errors.length) {
    throw new Error(
      'Đồng bộ trạng thái tài trợ Sheet 04 → Sheet 03 không khớp:\n- ' +
      errors.join('\n- ')
    );
  }
}

function FS03Sync_num_(value) {
  const n = Number(value);
  return isFinite(n) ? n : 0;
}
