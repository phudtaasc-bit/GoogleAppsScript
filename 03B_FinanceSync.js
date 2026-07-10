/*************************************************
 * 03B_FinanceSync.gs
 * Đồng bộ toàn bộ trạng thái tài trợ từ Sheet 04 về Sheet 03.
 * Sheet 04 là nguồn tính toán duy nhất cho cơ chế tài trợ.
 *************************************************/

function FS03_capNhatNguonVonTuSheet04_V2() {
  const ss = SpreadsheetApp.getActive();
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');

  if (!sh03 || !sh04) {
    throw new Error('Thiếu Sheet 03 hoặc Sheet 04 để đồng bộ nguồn vốn.');
  }

  const rows03 = Math.max(0, sh03.getLastRow() - 1);
  const rows04 = Math.max(0, sh04.getLastRow() - 2);
  if (!rows03 || !rows04) return;

  // Sheet 04: A:Z để lấy tháng, P, Q, S, V:Z.
  const data04 = sh04.getRange(3, 1, rows04, 26).getValues();
  const byMonth = {};

  data04.forEach(row => {
    const month = FS03V21_so_(row[0]);
    if (!month) return;

    byMonth[month] = {
      beforeFinancing: FS03V21_so_(row[15]), // P
      fundingNeed: FS03V21_so_(row[16]),     // Q
      equity: FS03V21_so_(row[18]),          // S
      loanDraw: FS03V21_so_(row[21]),        // V
      interest: FS03V21_so_(row[22]),        // W
      principal: FS03V21_so_(row[23]),       // X
      debtEnd: FS03V21_so_(row[24]),         // Y
      cashEnd: FS03V21_so_(row[25])          // Z
    };
  });

  const months03 = sh03.getRange(2, 1, rows03, 1).getValues();
  const output = [];
  let cumulativeEquity = 0;
  let cumulativeDraw = 0;
  let cumulativePrincipal = 0;

  months03.forEach(row => {
    const month = FS03V21_so_(row[0]);
    const v = byMonth[month] || {
      beforeFinancing: 0,
      fundingNeed: 0,
      equity: 0,
      loanDraw: 0,
      interest: 0,
      principal: 0,
      debtEnd: 0,
      cashEnd: 0
    };

    cumulativeEquity += v.equity;
    cumulativeDraw += v.loanDraw;
    cumulativePrincipal += v.principal;

    // AD là biến động tiền trong kỳ sau tài trợ và trả gốc;
    // AE là số dư tiền cuối kỳ chuyển sang kỳ sau.
    const cashMovementAfterFinancing =
      v.beforeFinancing + v.equity + v.loanDraw - v.principal;

    output.push([
      v.fundingNeed,                 // U  - Nhu cầu vốn
      v.equity,                      // V  - Vốn góp CSH
      cumulativeEquity,              // W  - Lũy kế vốn góp CSH
      v.loanDraw,                    // X  - Giải ngân vay
      cumulativeDraw,                // Y  - Lũy kế giải ngân vay
      v.interest,                    // Z  - Lãi vay vốn hóa
      v.principal,                   // AA - Trả gốc
      cumulativePrincipal,           // AB - Lũy kế trả gốc
      v.debtEnd,                     // AC - Dư nợ cuối kỳ
      cashMovementAfterFinancing,    // AD - Dòng tiền sau tài trợ
      v.cashEnd                      // AE - Tiền cuối kỳ
    ]);
  });

  sh03.getRange(2, 21, output.length, 11).setValues(output);
}
