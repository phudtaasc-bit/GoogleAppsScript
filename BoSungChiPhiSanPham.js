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
  relativeTolerance: 1e-8,
  maxRuntimeMs: 300000,          // Chủ động dừng ở phút thứ 5, trước giới hạn 6 phút.
  reserveForClosingMs: 45000,    // Dành thời gian đồng bộ và lập các sheet tổng hợp.
  stagnationLimit: 8,            // Số vòng liên tiếp gần như không cải thiện.
  minimumImprovementRatio: 0.001 // Cải thiện tối thiểu 0,1% mỗi vòng.
};

function FS_chayToanBoMoHinh_CoLaiVay() {
  const ss = SpreadsheetApp.getActive();
  const runStartedAt = Date.now();
  const profile = [];
  ss.toast('Đang chạy toàn bộ mô hình...', 'FS V2.1', 5);

  FS_profileStep_('Tạo Sheet 01 từ đầu vào', profile, () => FS_taoKyThuatTuDauVao());
  FS_profileStep_('Kiểm tra cấu hình tiền đất', profile, () => FS97_assertLandCostConfig_());
  FS_profileStep_('Lập Sheet 03 - chi phí/VAT', profile, () => FS_lapSheet03_CostOnly());
  FS_profileStep_('Lập Sheet 02 lần đầu', profile, () => FS_lapSheet02());
  FS_profileStep_('Lập Sheet 04 lần đầu', profile, () => FS_lapSheet04());
  SpreadsheetApp.flush();

  const result = FS_hoiTuLaiVay_(runStartedAt);
  if (!result.converged) {
    console.log(FS_moTaProfile_(profile, result, runStartedAt));
    throw new Error(FS_moTaLoiHoiTu_(result));
  }

  FS_profileStep_('Kiểm tra Thuế TNDN', profile, () => FS94_assertCITConsistency_());
  FS_profileStep_('Lập Sheet 04A', profile, () => FS_lapSheet04A());
  FS_profileStep_('Lập Sheet 00', profile, () => FS_lapSheet00());

  const profileText = FS_moTaProfile_(profile, result, runStartedAt);
  console.log(profileText);

  ss.toast(
    'Đã chạy xong trong ' + FS_formatSeconds_(Date.now() - runStartedAt) +
      '; tài trợ hội tụ sau ' + result.iterations + ' vòng. Chạy mục 7 để kiểm thử hồi quy.',
    'FS V2.1',
    10
  );
}

function FS_chayMoHinh_Buoc1() {
  const ss = SpreadsheetApp.getActive();
  const runStartedAt = Date.now();
  const profile = [];
  ss.toast('Đang chạy bước 1...', 'FS V2.1', 5);

  FS_profileStep_('Tạo Sheet 01 từ đầu vào', profile, () => FS_taoKyThuatTuDauVao());
  FS_profileStep_('Kiểm tra cấu hình tiền đất', profile, () => FS97_assertLandCostConfig_());
  FS_profileStep_('Lập Sheet 03 - chi phí/VAT', profile, () => FS_lapSheet03_CostOnly());
  FS_profileStep_('Lập Sheet 02 lần đầu', profile, () => FS_lapSheet02());
  FS_profileStep_('Lập Sheet 04 lần đầu', profile, () => FS_lapSheet04());
  SpreadsheetApp.flush();

  const result = FS_hoiTuLaiVay_(runStartedAt);
  if (!result.converged) {
    console.log(FS_moTaProfile_(profile, result, runStartedAt));
    throw new Error(FS_moTaLoiHoiTu_(result));
  }

  FS_profileStep_('Kiểm tra Thuế TNDN', profile, () => FS94_assertCITConsistency_());
  console.log(FS_moTaProfile_(profile, result, runStartedAt));

  ss.toast(
    'Xong bước 1 trong ' + FS_formatSeconds_(Date.now() - runStartedAt) +
      '; tài trợ hội tụ sau ' + result.iterations + ' vòng. Chạy tiếp bước 2.',
    'FS V2.1',
    10
  );
}

function FS_chayMoHinh_Buoc2() {
  const ss = SpreadsheetApp.getActive();
  const runStartedAt = Date.now();
  const profile = [];
  ss.toast('Đang chạy bước 2...', 'FS V2.1', 5);

  FS_profileStep_('Lập Sheet 02 lần đầu', profile, () => FS_lapSheet02());
  FS_profileStep_('Lập Sheet 04 lần đầu', profile, () => FS_lapSheet04());
  SpreadsheetApp.flush();

  const result = FS_hoiTuLaiVay_(runStartedAt);
  if (!result.converged) {
    console.log(FS_moTaProfile_(profile, result, runStartedAt));
    throw new Error(FS_moTaLoiHoiTu_(result));
  }

  FS_profileStep_('Kiểm tra Thuế TNDN', profile, () => FS94_assertCITConsistency_());
  FS_profileStep_('Lập Sheet 04A', profile, () => FS_lapSheet04A());
  FS_profileStep_('Lập Sheet 00', profile, () => FS_lapSheet00());
  console.log(FS_moTaProfile_(profile, result, runStartedAt));

  ss.toast(
    'Đã chạy xong bước 2 trong ' + FS_formatSeconds_(Date.now() - runStartedAt) +
      '. Chạy mục 7 để kiểm thử hồi quy.',
    'FS V2.1',
    10
  );
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
function FS_hoiTuLaiVay_(runStartedAt) {
  const startedAt = runStartedAt || Date.now();
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
  let previousDelta = Infinity;
  let stagnationCount = 0;
  const iterationProfile = [];

  for (let iteration = 1; iteration <= FS_FINANCE_ITERATION.maxIterations; iteration++) {
    const elapsedBefore = Date.now() - startedAt;
    const safeDeadline = FS_FINANCE_ITERATION.maxRuntimeMs - FS_FINANCE_ITERATION.reserveForClosingMs;
    if (elapsedBefore >= safeDeadline) {
      return Object.assign({
        converged: false,
        iterations: iteration - 1,
        tolerance: tolerance,
        stoppedByRuntime: true,
        elapsedMs: elapsedBefore,
        iterationProfile: iterationProfile
      }, comparison);
    }

    const iterationStartedAt = Date.now();
    const stepTimes = {};

    let t = Date.now();
    FS03_capNhatNguonVonTuSheet04_V2();
    stepTimes.sync03Ms = Date.now() - t;

    t = Date.now();
    FS_lapSheet02();
    stepTimes.sheet02Ms = Date.now() - t;

    t = Date.now();
    FS_lapSheet04();
    stepTimes.sheet04Ms = Date.now() - t;

    t = Date.now();
    SpreadsheetApp.flush();
    stepTimes.flushMs = Date.now() - t;

    const current = FS_docTrangThaiTaiTro_();
    comparison = FS_saiLechTrangThaiTaiTro_(previous, current);
    const scale = Math.max(1, current.maxAbsoluteValue);
    tolerance = Math.max(
      FS_FINANCE_ITERATION.absoluteTolerance,
      scale * FS_FINANCE_ITERATION.relativeTolerance
    );

    const improvementRatio = isFinite(previousDelta) && previousDelta > 0
      ? (previousDelta - comparison.maxDelta) / previousDelta
      : 1;

    if (comparison.maxDelta > tolerance && improvementRatio < FS_FINANCE_ITERATION.minimumImprovementRatio) {
      stagnationCount++;
    } else {
      stagnationCount = 0;
    }

    const item = {
      iteration: iteration,
      elapsedMs: Date.now() - startedAt,
      iterationMs: Date.now() - iterationStartedAt,
      sync03Ms: stepTimes.sync03Ms,
      sheet02Ms: stepTimes.sheet02Ms,
      sheet04Ms: stepTimes.sheet04Ms,
      flushMs: stepTimes.flushMs,
      maxDelta: comparison.maxDelta,
      tolerance: tolerance,
      improvementRatio: improvementRatio
    };
    iterationProfile.push(item);

    console.log(
      '[FS hội tụ] vòng %s | %ss | Δmax=%s | ngưỡng=%s | cải thiện=%s%% | sync03=%ss | S02=%ss | S04=%ss | flush=%ss',
      iteration,
      FS_formatSeconds_(item.iterationMs),
      Math.round(comparison.maxDelta).toLocaleString('vi-VN'),
      Math.round(tolerance).toLocaleString('vi-VN'),
      (improvementRatio * 100).toFixed(3),
      FS_formatSeconds_(stepTimes.sync03Ms),
      FS_formatSeconds_(stepTimes.sheet02Ms),
      FS_formatSeconds_(stepTimes.sheet04Ms),
      FS_formatSeconds_(stepTimes.flushMs)
    );

    if (comparison.maxDelta <= tolerance) {
      FS03_capNhatNguonVonTuSheet04_V2();
      SpreadsheetApp.flush();
      return Object.assign({
        converged: true,
        iterations: iteration,
        tolerance: tolerance,
        elapsedMs: Date.now() - startedAt,
        iterationProfile: iterationProfile
      }, comparison);
    }

    if (stagnationCount >= FS_FINANCE_ITERATION.stagnationLimit) {
      return Object.assign({
        converged: false,
        iterations: iteration,
        tolerance: tolerance,
        stoppedByStagnation: true,
        stagnationCount: stagnationCount,
        elapsedMs: Date.now() - startedAt,
        iterationProfile: iterationProfile
      }, comparison);
    }

    previous = current;
    previousDelta = comparison.maxDelta;
  }

  return Object.assign({
    converged: false,
    iterations: FS_FINANCE_ITERATION.maxIterations,
    tolerance: tolerance,
    elapsedMs: Date.now() - startedAt,
    iterationProfile: iterationProfile
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
  const reason = result.stoppedByRuntime
    ? 'Đã chủ động dừng trước giới hạn 6 phút của Apps Script.'
    : result.stoppedByStagnation
      ? 'Sai số gần như không giảm trong ' + result.stagnationCount + ' vòng liên tiếp.'
      : 'Đã đạt giới hạn số vòng lặp.';

  return [
    'Mô hình tài trợ chưa hội tụ sau ' + result.iterations + ' vòng.',
    reason,
    'Thời gian đã chạy: ' + FS_formatSeconds_(result.elapsedMs || 0) + '.',
    'Sai lệch lớn nhất: ' + fmt(result.maxDelta) + ' đồng; ngưỡng: ' + fmt(result.tolerance) + ' đồng.',
    'Giải ngân: ' + fmt(result.drawDelta) + '; ' +
      'Lãi vay: ' + fmt(result.interestDelta) + '; ' +
      'Trả gốc: ' + fmt(result.principalDelta) + '; ' +
      'Dư nợ: ' + fmt(result.debtDelta) + '; ' +
      'Tiền cuối kỳ: ' + fmt(result.cashDelta) + ' đồng.',
    'Mở Nhật ký thực thi để xem thời gian Sheet 02, Sheet 04 và từng vòng hội tụ.'
  ].join('\n');
}

function FS_profileStep_(name, profile, action) {
  const startedAt = Date.now();
  const result = action();
  const elapsedMs = Date.now() - startedAt;
  profile.push({ name: name, elapsedMs: elapsedMs });
  console.log('[FS thời gian] %s: %s', name, FS_formatSeconds_(elapsedMs));
  return result;
}

function FS_moTaProfile_(profile, result, runStartedAt) {
  const lines = ['=== FS PERFORMANCE PROFILE ==='];
  (profile || []).forEach(x => lines.push(x.name + ': ' + FS_formatSeconds_(x.elapsedMs)));
  lines.push('Hội tụ: ' + (result.converged ? 'ĐẠT' : 'CHƯA ĐẠT'));
  lines.push('Số vòng: ' + (result.iterations || 0));
  lines.push('Thời gian hội tụ/lũy kế: ' + FS_formatSeconds_(result.elapsedMs || 0));
  lines.push('Tổng thời gian: ' + FS_formatSeconds_(Date.now() - runStartedAt));
  lines.push('Sai lệch cuối: ' + Math.round(Number(result.maxDelta) || 0).toLocaleString('vi-VN') + ' đồng');
  return lines.join('\n');
}

function FS_formatSeconds_(milliseconds) {
  return ((Number(milliseconds) || 0) / 1000).toFixed(2) + ' giây';
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
    'Kiểm thử hồi quy chạy riêng tại mục 7 để tránh vượt giới hạn thời gian Apps Script.',
    'Runner ghi thời gian từng bước và chủ động dừng trước giới hạn 6 phút.',
    'Số vòng tối đa: ' + FS_FINANCE_ITERATION.maxIterations + '.',
    'Sai số tuyệt đối tối thiểu: ' + FS_FINANCE_ITERATION.absoluteTolerance.toLocaleString('vi-VN') + ' đồng.'
  ].join('\n');

  SpreadsheetApp.getUi().alert(msg);
}