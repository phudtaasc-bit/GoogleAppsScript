/*************************************************
 * FS V2.1 - MODEL RUNNER
 *
 * Giữ nguyên các hàm lập Sheet 00/02/03/04 hiện có.
 * Luồng chạy sử dụng Sheet 03 đã vá và lặp Sheet 02-04
 * đến khi lãi vay, dư nợ hội tụ.
 *************************************************/

const FS_FINANCE_ITERATION = {
  maxIterations: 20,
  absoluteTolerance: 1000,
  relativeTolerance: 1e-8
};

function FS_chayToanBoMoHinh_CoLaiVay() {
  const ss = SpreadsheetApp.getActive();
  ss.toast('Đang chạy toàn bộ mô hình...', 'FS V2.1', 5);

  FS_taoKyThuatTuDauVao();
  FS97_assertLandCostConfig_();

  FS_lapSheet03_Patched();
  FS_lapSheet02();
  FS_lapSheet04();
  SpreadsheetApp.flush();

  const result = FS_hoiTuLaiVay_();
  if (!result.converged) {
    throw new Error(
      'Mô hình lãi vay chưa hội tụ sau ' + result.iterations + ' vòng. ' +
      'Sai lệch lớn nhất còn lại: ' + Math.round(result.maxDelta).toLocaleString('vi-VN') + ' đồng. ' +
      'Dừng trước khi lập các sheet tổng hợp để tránh sử dụng kết quả chưa ổn định.'
    );
  }

  FS94_assertCITConsistency_();
  FS_lapSheet04A();
  FS_lapSheet00_Patched();
  FS93_runRegressionSuite_({ showAlert: false, throwOnError: true });

  ss.toast(
    'Đã chạy xong toàn bộ mô hình; lãi vay hội tụ sau ' + result.iterations + ' vòng và kiểm thử hồi quy đạt.',
    'FS V2.1',
    10
  );
}

function FS_chayMoHinh_Buoc1() {
  const ss = SpreadsheetApp.getActive();
  ss.toast('Đang chạy bước 1...', 'FS V2.1', 5);

  FS_taoKyThuatTuDauVao();
  FS97_assertLandCostConfig_();

  FS_lapSheet03_Patched();
  FS_lapSheet02();
  FS_lapSheet04();
  SpreadsheetApp.flush();

  const result = FS_hoiTuLaiVay_();
  if (!result.converged) {
    throw new Error(
      'Mô hình lãi vay chưa hội tụ sau ' + result.iterations + ' vòng. ' +
      'Sai lệch lớn nhất còn lại: ' + Math.round(result.maxDelta).toLocaleString('vi-VN') + ' đồng.'
    );
  }

  FS94_assertCITConsistency_();

  ss.toast(
    'Xong bước 1; lãi vay hội tụ sau ' + result.iterations + ' vòng. Chạy tiếp bước 2.',
    'FS V2.1',
    10
  );
}

function FS_chayMoHinh_Buoc2() {
  const ss = SpreadsheetApp.getActive();
  ss.toast('Đang chạy bước 2...', 'FS V2.1', 5);

  FS_lapSheet02();
  FS_lapSheet04();
  SpreadsheetApp.flush();

  const result = FS_hoiTuLaiVay_();
  if (!result.converged) {
    throw new Error(
      'Mô hình lãi vay chưa hội tụ sau ' + result.iterations + ' vòng. ' +
      'Sai lệch lớn nhất còn lại: ' + Math.round(result.maxDelta).toLocaleString('vi-VN') + ' đồng.'
    );
  }

  FS94_assertCITConsistency_();
  FS_lapSheet04A();
  FS_lapSheet00_Patched();
  FS93_runRegressionSuite_({ showAlert: false, throwOnError: true });

  ss.toast('Đã chạy xong bước 2 và kiểm thử hồi quy đạt.', 'FS V2.1', 10);
}

/**
 * Lặp cố định giữa:
 * - Sheet 04: xác định lãi vay và dư nợ theo dòng tiền;
 * - Sheet 03: nhận lại nguồn vốn/lãi vay;
 * - Sheet 02: phân bổ lãi vay vào giá vốn và tính Thuế TNDN;
 * - Sheet 04: tính lại dòng tiền, nhu cầu vốn và dư nợ.
 */
function FS_hoiTuLaiVay_() {
  let previous = FS_docTrangThaiTaiTro_();
  let maxDelta = Infinity;

  for (let iteration = 1; iteration <= FS_FINANCE_ITERATION.maxIterations; iteration++) {
    FS03_capNhatNguonVonTuSheet04();
    FS_lapSheet02();
    FS_lapSheet04();
    SpreadsheetApp.flush();

    const current = FS_docTrangThaiTaiTro_();
    maxDelta = FS_saiLechTrangThaiTaiTro_(previous, current);
    const scale = Math.max(1, current.maxAbsoluteValue);
    const tolerance = Math.max(
      FS_FINANCE_ITERATION.absoluteTolerance,
      scale * FS_FINANCE_ITERATION.relativeTolerance
    );

    if (maxDelta <= tolerance) {
      // Đồng bộ Sheet 03 lần cuối với trạng thái Sheet 04 đã hội tụ.
      FS03_capNhatNguonVonTuSheet04();
      SpreadsheetApp.flush();
      return { converged: true, iterations: iteration, maxDelta, tolerance };
    }

    previous = current;
  }

  return {
    converged: false,
    iterations: FS_FINANCE_ITERATION.maxIterations,
    maxDelta,
    tolerance: FS_FINANCE_ITERATION.absoluteTolerance
  };
}

function FS_docTrangThaiTaiTro_() {
  const ss = SpreadsheetApp.getActive();
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');
  if (!sh04 || sh04.getLastRow() < 3) {
    throw new Error('Sheet 04 chưa có dữ liệu để kiểm tra hội tụ lãi vay.');
  }

  const numRows = sh04.getLastRow() - 2;
  const values = sh04.getRange(3, 22, numRows, 4).getValues(); // V:Y
  const vector = [];
  let maxAbsoluteValue = 0;

  values.forEach(r => {
    // Theo dõi giải ngân, lãi vay, trả gốc và dư nợ để tránh hội tụ giả.
    r.forEach(v => {
      const n = Number(v);
      const value = isFinite(n) ? n : 0;
      vector.push(value);
      maxAbsoluteValue = Math.max(maxAbsoluteValue, Math.abs(value));
    });
  });

  return { vector, maxAbsoluteValue };
}

function FS_saiLechTrangThaiTaiTro_(a, b) {
  const n = Math.max(a.vector.length, b.vector.length);
  let maxDelta = 0;
  for (let i = 0; i < n; i++) {
    const av = Number(a.vector[i]) || 0;
    const bv = Number(b.vector[i]) || 0;
    maxDelta = Math.max(maxDelta, Math.abs(av - bv));
  }
  return maxDelta;
}

/**
 * Chạy kiểm tra nhanh để phát hiện project còn patch cũ hay không.
 */
function FS_V21_KiemTraNhanhSauKhiDan() {
  const ss = SpreadsheetApp.getActive();
  const msg = [
    'Đã nạp runner FS V2.1 có kiểm tra hội tụ lãi vay.',
    'Luồng chuẩn sử dụng FS_lapSheet03_Patched.',
    'Vòng lặp dừng khi giải ngân, lãi vay, trả gốc và dư nợ ổn định.',
    'Thuế TNDN được kiểm tra riêng theo từng sản phẩm trước khi lập sheet tổng hợp.',
    'NPV dự án dùng WACC; NPV vốn CSH dùng chi phí vốn chủ sở hữu.',
    'Kiểm thử hồi quy được chạy sau khi lập Sheet 00.',
    'Số vòng tối đa: ' + FS_FINANCE_ITERATION.maxIterations + '.',
    'Sai số tuyệt đối tối thiểu: ' + FS_FINANCE_ITERATION.absoluteTolerance.toLocaleString('vi-VN') + ' đồng.'
  ].join('\n');

  SpreadsheetApp.getUi().alert(msg);
}
