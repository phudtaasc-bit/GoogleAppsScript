/*************************************************
 * FS V2.1 - MODEL RUNNER
 *
 * Giữ nguyên các hàm lập Sheet 00/02/03/04 hiện có.
 * Luồng chạy sử dụng Sheet 03 chỉ tính chi phí/VAT và lặp Sheet 02-04
 * đến khi toàn bộ trạng thái tài trợ hội tụ.
 *************************************************/

const FS_FINANCE_ITERATION = {
  maxIterations: 50,
  absoluteTolerance: 1000,
  relativeTolerance: 1e-8
};

function FS_chayToanBoMoHinh_CoLaiVay() {
  const ss = SpreadsheetApp.getActive();
  ss.toast('Đang chạy toàn bộ mô hình...', 'FS V2.1', 5);

  FS_taoKyThuatTuDauVao();
  FS97_assertLandCostConfig_();

  FS_lapSheet03_CostOnly();
  FS_lapSheet02();
  FS_lapSheet04();
  SpreadsheetApp.flush();

  const result = FS_hoiTuLaiVay_();
  if (!result.converged) {
    throw new Error(FS_moTaLoiHoiTu_(result));
  }

  FS94_assertCITConsistency_();
  FS_lapSheet04A();
  FS_lapSheet00();
  FS93_runRegressionSuite_({ showAlert: false, throwOnError: true });

  ss.toast(
    'Đã chạy xong toàn bộ mô hình; tài trợ hội tụ sau ' + result.iterations + ' vòng và kiểm thử hồi quy đạt.',
    'FS V2.1',
    10
  );
}

function FS_chayMoHinh_Buoc1() {
  const ss = SpreadsheetApp.getActive();
  ss.toast('Đang chạy bước 1...', 'FS V2.1', 5);

  FS_taoKyThuatTuDauVao();
  FS97_assertLandCostConfig_();

  FS_lapSheet03_CostOnly();
  FS_lapSheet02();
  FS_lapSheet04();
  SpreadsheetApp.flush();

  const result = FS_hoiTuLaiVay_();
  if (!result.converged) {
    throw new Error(FS_moTaLoiHoiTu_(result));
  }

  FS94_assertCITConsistency_();

  ss.toast(
    'Xong bước 1; tài trợ hội tụ sau ' + result.iterations + ' vòng. Chạy tiếp bước 2.',
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
    throw new Error(FS_moTaLoiHoiTu_(result));
  }

  FS94_assertCITConsistency_();
  FS_lapSheet04A();
  FS_lapSheet00();
  FS93_runRegressionSuite_({ showAlert: false, throwOnError: true });

  ss.toast('Đã chạy xong bước 2 và kiểm thử hồi quy đạt.', 'FS V2.1', 10);
}

/**
 * Lặp cố định giữa:
 * - Sheet 04: engine xác định tài trợ, lãi vay vốn hóa, trả gốc và tiền giữ lại;
 * - Sheet 03: nhận lại toàn bộ trạng thái tài trợ để phân bổ chi phí lãi vay;
 * - Sheet 02: phân bổ lãi vay vào giá vốn và tính Thuế TNDN;
 * - Sheet 04: tính lại cash waterfall.
 *
 * Hội tụ chỉ được công nhận khi đồng thời ổn định:
 * giải ngân, lãi vay, trả gốc, dư nợ và tiền cuối kỳ.
 */
function FS_hoiTuLaiVay_() {
  let previous = FS_docTrangThaiTaiTro_();
  let comparison = {
    maxDelta: Infinity,
    drawDelta: Infinity,
    interestDelta: Infinity,
    principalDelta: Infinity,
    debtDelta: Infinity,
    cashDelta: Infinity
  };
  let tolerance = FS_FINANCE_ITERATION.absoluteTolerance;

  for (let iteration = 1; iteration <= FS_FINANCE_ITERATION.maxIterations; iteration++) {
    FS03_capNhatNguonVonTuSheet04_V2();
    FS_lapSheet02();
    FS_lapSheet04();
    SpreadsheetApp.flush();

    const current = FS_docTrangThaiTaiTro_();
    comparison = FS_saiLechTrangThaiTaiTro_(previous, current);
    const scale = Math.max(1, current.maxAbsoluteValue);
    tolerance = Math.max(
      FS_FINANCE_ITERATION.absoluteTolerance,
      scale * FS_FINANCE_ITERATION.relativeTolerance
    );

    if (comparison.maxDelta <= tolerance) {
      FS03_capNhatNguonVonTuSheet04_V2();
      SpreadsheetApp.flush();
      return Object.assign({
        converged: true,
        iterations: iteration,
        tolerance: tolerance
      }, comparison);
    }

    previous = current;
  }

  return Object.assign({
    converged: false,
    iterations: FS_FINANCE_ITERATION.maxIterations,
    tolerance: tolerance
  }, comparison);
}

function FS_docTrangThaiTaiTro_() {
  const ss = SpreadsheetApp.getActive();
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');
  if (!sh04 || sh04.getLastRow() < 3) {
    throw new Error('Sheet 04 chưa có dữ liệu để kiểm tra hội tụ tài trợ.');
  }

  const numRows = sh04.getLastRow() - 2;
  const values = sh04.getRange(3, 22, numRows, 5).getValues(); // V:Z
  const vector = [];
  const byMetric = [[], [], [], [], []];
  let maxAbsoluteValue = 0;

  values.forEach(row => {
    row.forEach((v, metricIndex) => {
      const n = Number(v);
      const value = isFinite(n) ? n : 0;
      vector.push(value);
      byMetric[metricIndex].push(value);
      maxAbsoluteValue = Math.max(maxAbsoluteValue, Math.abs(value));
    });
  });

  return {
    vector: vector,
    byMetric: byMetric,
    maxAbsoluteValue: maxAbsoluteValue
  };
}

function FS_saiLechTrangThaiTaiTro_(a, b) {
  const metricNames = [
    'drawDelta',
    'interestDelta',
    'principalDelta',
    'debtDelta',
    'cashDelta'
  ];
  const result = { maxDelta: 0 };

  metricNames.forEach((name, metricIndex) => {
    const av = (a.byMetric && a.byMetric[metricIndex]) || [];
    const bv = (b.byMetric && b.byMetric[metricIndex]) || [];
    const n = Math.max(av.length, bv.length);
    let metricDelta = 0;

    for (let i = 0; i < n; i++) {
      metricDelta = Math.max(
        metricDelta,
        Math.abs((Number(av[i]) || 0) - (Number(bv[i]) || 0))
      );
    }

    result[name] = metricDelta;
    result.maxDelta = Math.max(result.maxDelta, metricDelta);
  });

  return result;
}

function FS_moTaLoiHoiTu_(result) {
  const fmt = value => Math.round(Number(value) || 0).toLocaleString('vi-VN');
  return [
    'Mô hình tài trợ chưa hội tụ sau ' + result.iterations + ' vòng.',
    'Sai lệch lớn nhất: ' + fmt(result.maxDelta) + ' đồng.',
    'Giải ngân: ' + fmt(result.drawDelta) + '; ' +
      'Lãi vay: ' + fmt(result.interestDelta) + '; ' +
      'Trả gốc: ' + fmt(result.principalDelta) + '; ' +
      'Dư nợ: ' + fmt(result.debtDelta) + '; ' +
      'Tiền cuối kỳ: ' + fmt(result.cashDelta) + ' đồng.',
    'Dừng trước khi lập các sheet tổng hợp để tránh sử dụng kết quả chưa ổn định.'
  ].join('\n');
}

function FS_V21_KiemTraNhanhSauKhiDan() {
  const msg = [
    'Đã nạp runner FS V2.1 có kiểm tra hội tụ tài trợ.',
    'Luồng chuẩn khởi tạo Sheet 03 ở chế độ chỉ tính chi phí/VAT.',
    'Cơ cấu vốn áp dụng theo từng lần thiếu vốn (Cách A).',
    'Sheet 04 là nguồn tính toán duy nhất; Sheet 03 nhận lại toàn bộ khối tài trợ U:AE.',
    'Vòng lặp dừng khi giải ngân, lãi vay, trả gốc, dư nợ và tiền cuối kỳ cùng ổn định.',
    'Lãi vay được vốn hóa vào dư nợ.',
    'Thuế TNDN được kiểm tra riêng theo từng sản phẩm trước khi lập sheet tổng hợp.',
    'NPV dự án dùng WACC; NPV vốn CSH dùng chi phí vốn chủ sở hữu.',
    'Kiểm thử hồi quy được chạy sau khi lập Sheet 00.',
    'Số vòng tối đa: ' + FS_FINANCE_ITERATION.maxIterations + '.',
    'Sai số tuyệt đối tối thiểu: ' + FS_FINANCE_ITERATION.absoluteTolerance.toLocaleString('vi-VN') + ' đồng.'
  ].join('\n');

  SpreadsheetApp.getUi().alert(msg);
}
