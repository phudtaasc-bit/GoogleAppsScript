/*************************************************
 * FS V2.2 - MODEL RUNNER PHÂN ĐOẠN
 *
 * Giữ nguyên toàn bộ logic tài chính và các hàm lập Sheet hiện có.
 * Chỉ tách quá trình thực thi thành nhiều lần chạy để không vượt 6 phút.
 *************************************************/

const FS_FINANCE_ITERATION = {
  maxIterations: 50,
  maxIterationsPerRun: 2,
  absoluteTolerance: 1000,
  relativeTolerance: 1e-8,
  maxRuntimeMs: 300000,
  reserveForClosingMs: 45000,
  stagnationLimit: 8,
  minimumImprovementRatio: 0.001
};

const FS_RUN_STATE = {
  prefix: 'FS_V22_',
  status: 'STATUS',
  totalIterations: 'TOTAL_ITERATIONS',
  lastDelta: 'LAST_DELTA',
  lastTolerance: 'LAST_TOLERANCE',
  stagnationCount: 'STAGNATION_COUNT',
  updatedAt: 'UPDATED_AT'
};

/**
 * Hàm tương thích với menu 4.3 hiện tại.
 * Mỗi lần bấm sẽ tự chạy bước phù hợp theo trạng thái đã lưu:
 * - Chưa khởi tạo -> Bước 1
 * - Đang hội tụ -> Bước 2
 * - Đã hội tụ -> Bước 3
 * - Đã hoàn tất -> thông báo
 */
function FS_chayToanBoMoHinh_CoLaiVay() {
  const state = FS_getRunState_();

  if (!state.status) {
    FS_chayMoHinh_Buoc1();
    return;
  }

  if (state.status === 'INITIALIZED' || state.status === 'ITERATING') {
    FS_chayMoHinh_Buoc2();
    return;
  }

  if (state.status === 'CONVERGED') {
    FS_chayMoHinh_Buoc3();
    return;
  }

  if (state.status === 'FINALIZED') {
    SpreadsheetApp.getUi().alert(
      'Mô hình đã hoàn tất.\n' +
      'Muốn tính lại từ đầu, chạy chức năng "Xóa trạng thái chạy" rồi chạy lại Bước 1.'
    );
    return;
  }

  throw new Error('Trạng thái chạy mô hình không hợp lệ: ' + state.status);
}

/**
 * BƯỚC 1 - KHỞI TẠO
 * Chỉ lập các sheet nền và lưu trạng thái. Không chạy hội tụ.
 */
function FS_chayMoHinh_Buoc1() {
  const ss = SpreadsheetApp.getActive();
  const runStartedAt = Date.now();
  const profile = [];

  ss.toast('Đang khởi tạo mô hình...', 'FS V2.2 - Bước 1', 5);
  FS_xoaTrangThaiChayNoAlert_();

  FS_profileStep_('Tạo Sheet 01 từ đầu vào', profile, () => FS_taoKyThuatTuDauVao());
  FS_profileStep_('Kiểm tra cấu hình tiền đất', profile, () => FS97_assertLandCostConfig_());
  FS_profileStep_('Lập Sheet 03 - chi phí/VAT', profile, () => FS_lapSheet03_CostOnly());
  FS_profileStep_('Lập Sheet 02 lần đầu', profile, () => FS_lapSheet02());
  FS_profileStep_('Lập Sheet 04 lần đầu', profile, () => FS_lapSheet04());
  SpreadsheetApp.flush();

  FS_setRunState_({
    status: 'INITIALIZED',
    totalIterations: 0,
    lastDelta: '',
    lastTolerance: '',
    stagnationCount: 0
  });

  console.log(FS_moTaProfile_(profile, {
    converged: false,
    iterations: 0,
    elapsedMs: Date.now() - runStartedAt,
    maxDelta: 0
  }, runStartedAt));

  ss.toast(
    'Đã xong Bước 1 trong ' + FS_formatSeconds_(Date.now() - runStartedAt) +
    '. Chạy tiếp Bước 2 để hội tụ tài trợ.',
    'FS V2.2',
    10
  );
}

/**
 * BƯỚC 2 - HỘI TỤ TÀI TRỢ
 * Mỗi lần chạy tối đa số vòng quy định trong maxIterationsPerRun.
 * Nếu chưa hội tụ, lưu trạng thái và tiếp tục ở lần chạy sau.
 */
function FS_chayMoHinh_Buoc2() {
  const ss = SpreadsheetApp.getActive();
  const runStartedAt = Date.now();
  const state = FS_getRunState_();

  if (!state.status) {
    throw new Error('Chưa khởi tạo mô hình. Hãy chạy Bước 1 trước.');
  }
  if (state.status === 'CONVERGED') {
    ss.toast('Tài trợ đã hội tụ. Chạy Bước 3 để hoàn tất báo cáo.', 'FS V2.2', 10);
    return;
  }
  if (state.status === 'FINALIZED') {
    ss.toast('Mô hình đã hoàn tất.', 'FS V2.2', 10);
    return;
  }
  if (state.status !== 'INITIALIZED' && state.status !== 'ITERATING') {
    throw new Error('Không thể tiếp tục hội tụ từ trạng thái: ' + state.status);
  }

  ss.toast(
    'Đang hội tụ tài trợ từ vòng ' + (state.totalIterations + 1) + '...',
    'FS V2.2 - Bước 2',
    5
  );

  const result = FS_hoiTuLaiVay_PhanDoan_(runStartedAt, state);

  FS_setRunState_({
    status: result.converged ? 'CONVERGED' : 'ITERATING',
    totalIterations: result.totalIterations,
    lastDelta: result.maxDelta,
    lastTolerance: result.tolerance,
    stagnationCount: result.stagnationCount || 0
  });

  console.log(FS_moTaProfile_([], result, runStartedAt));

  if (result.converged) {
    ss.toast(
      'Tài trợ đã hội tụ sau tổng cộng ' + result.totalIterations +
      ' vòng. Chạy Bước 3 để hoàn tất báo cáo.',
      'FS V2.2',
      10
    );
    return;
  }

  if (result.stoppedByStagnation) {
    throw new Error(FS_moTaLoiHoiTu_(result));
  }

  ss.toast(
    'Đã chạy thêm ' + result.iterationsThisRun + ' vòng; tổng cộng ' +
    result.totalIterations + ' vòng. Chưa hội tụ, hãy chạy lại Bước 2.',
    'FS V2.2',
    10
  );
}

/**
 * BƯỚC 3 - HOÀN TẤT
 * Chỉ được chạy sau khi Bước 2 xác nhận hội tụ.
 */
function FS_chayMoHinh_Buoc3() {
  const ss = SpreadsheetApp.getActive();
  const runStartedAt = Date.now();
  const profile = [];
  const state = FS_getRunState_();

  if (state.status !== 'CONVERGED') {
    if (state.status === 'FINALIZED') {
      ss.toast('Mô hình đã hoàn tất.', 'FS V2.2', 10);
      return;
    }
    throw new Error('Tài trợ chưa hội tụ. Hãy chạy Bước 2 cho đến khi có thông báo hội tụ.');
  }

  ss.toast('Đang hoàn tất mô hình...', 'FS V2.2 - Bước 3', 5);

  FS_profileStep_('Đồng bộ nguồn vốn lần cuối', profile, () => FS03_capNhatNguonVonTuSheet04_V2());
  SpreadsheetApp.flush();
  FS_profileStep_('Kiểm tra Thuế TNDN', profile, () => FS94_assertCITConsistency_());
  FS_profileStep_('Lập Sheet 04A', profile, () => FS_lapSheet04A());
  FS_profileStep_('Lập Sheet 00', profile, () => FS_lapSheet00());

  FS_setRunState_({
    status: 'FINALIZED',
    totalIterations: state.totalIterations,
    lastDelta: state.lastDelta,
    lastTolerance: state.lastTolerance,
    stagnationCount: state.stagnationCount
  });

  console.log(FS_moTaProfile_(profile, {
    converged: true,
    iterations: state.totalIterations,
    totalIterations: state.totalIterations,
    elapsedMs: Date.now() - runStartedAt,
    maxDelta: state.lastDelta
  }, runStartedAt));

  ss.toast(
    'Đã hoàn tất mô hình trong ' + FS_formatSeconds_(Date.now() - runStartedAt) +
    '. Tổng số vòng hội tụ: ' + state.totalIterations + '.',
    'FS V2.2',
    10
  );
}

/**
 * Xóa trạng thái tiến trình, không xóa dữ liệu trên Sheet.
 */
function FS_xoaTrangThaiChay() {
  FS_xoaTrangThaiChayNoAlert_();
  SpreadsheetApp.getUi().alert(
    'Đã xóa trạng thái chạy mô hình.\nDữ liệu trên các Sheet không bị xóa.\nChạy Bước 1 để tính lại từ đầu.'
  );
}

function FS_xoaTrangThaiChayNoAlert_() {
  const props = PropertiesService.getDocumentProperties();
  Object.keys(FS_RUN_STATE).forEach(key => {
    if (key !== 'prefix') {
      props.deleteProperty(FS_RUN_STATE.prefix + FS_RUN_STATE[key]);
    }
  });
}

function FS_getRunState_() {
  const props = PropertiesService.getDocumentProperties();
  const get = key => props.getProperty(FS_RUN_STATE.prefix + FS_RUN_STATE[key]);

  return {
    status: get('status') || '',
    totalIterations: Number(get('totalIterations')) || 0,
    lastDelta: Number(get('lastDelta')) || 0,
    lastTolerance: Number(get('lastTolerance')) || FS_FINANCE_ITERATION.absoluteTolerance,
    stagnationCount: Number(get('stagnationCount')) || 0,
    updatedAt: get('updatedAt') || ''
  };
}

function FS_setRunState_(state) {
  const props = PropertiesService.getDocumentProperties();
  const values = {};

  values[FS_RUN_STATE.prefix + FS_RUN_STATE.status] = String(state.status || '');
  values[FS_RUN_STATE.prefix + FS_RUN_STATE.totalIterations] = String(Number(state.totalIterations) || 0);
  values[FS_RUN_STATE.prefix + FS_RUN_STATE.lastDelta] =
    state.lastDelta === '' ? '' : String(Number(state.lastDelta) || 0);
  values[FS_RUN_STATE.prefix + FS_RUN_STATE.lastTolerance] =
    state.lastTolerance === '' ? '' : String(Number(state.lastTolerance) || 0);
  values[FS_RUN_STATE.prefix + FS_RUN_STATE.stagnationCount] =
    String(Number(state.stagnationCount) || 0);
  values[FS_RUN_STATE.prefix + FS_RUN_STATE.updatedAt] = new Date().toISOString();

  props.setProperties(values, false);
}

/**
 * Chạy một phân đoạn hội tụ và trả lại trạng thái để lưu cho lần sau.
 */
function FS_hoiTuLaiVay_PhanDoan_(runStartedAt, savedState) {
  const startedAt = runStartedAt || Date.now();
  let previous = FS_docTrangThaiTaiTro_();
  let comparison = {
    maxDelta: Number(savedState.lastDelta) || Infinity,
    drawDelta: Infinity,
    interestDelta: Infinity,
    principalDelta: Infinity,
    debtDelta: Infinity,
    cashDelta: Infinity
  };
  let tolerance = Number(savedState.lastTolerance) || FS_FINANCE_ITERATION.absoluteTolerance;
  let previousDelta = Number(savedState.lastDelta) || Infinity;
  let stagnationCount = Number(savedState.stagnationCount) || 0;
  let totalIterations = Number(savedState.totalIterations) || 0;
  const iterationProfile = [];

  for (
    let localIteration = 1;
    localIteration <= FS_FINANCE_ITERATION.maxIterationsPerRun;
    localIteration++
  ) {
    if (totalIterations >= FS_FINANCE_ITERATION.maxIterations) {
      return Object.assign({
        converged: false,
        iterations: totalIterations,
        iterationsThisRun: localIteration - 1,
        totalIterations: totalIterations,
        tolerance: tolerance,
        elapsedMs: Date.now() - startedAt,
        reachedMaxIterations: true,
        stagnationCount: stagnationCount,
        iterationProfile: iterationProfile
      }, comparison);
    }

    const elapsedBefore = Date.now() - startedAt;
    const safeDeadline = FS_FINANCE_ITERATION.maxRuntimeMs - FS_FINANCE_ITERATION.reserveForClosingMs;
    if (elapsedBefore >= safeDeadline) {
      return Object.assign({
        converged: false,
        iterations: totalIterations,
        iterationsThisRun: localIteration - 1,
        totalIterations: totalIterations,
        tolerance: tolerance,
        stoppedByRuntime: true,
        elapsedMs: elapsedBefore,
        stagnationCount: stagnationCount,
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

    if (
      comparison.maxDelta > tolerance &&
      improvementRatio < FS_FINANCE_ITERATION.minimumImprovementRatio
    ) {
      stagnationCount++;
    } else {
      stagnationCount = 0;
    }

    totalIterations++;

    const item = {
      iteration: totalIterations,
      localIteration: localIteration,
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
      '[FS hội tụ] vòng tổng %s | vòng phiên %s | %s | Δmax=%s | ngưỡng=%s | cải thiện=%s%% | sync03=%s | S02=%s | S04=%s | flush=%s',
      totalIterations,
      localIteration,
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
        iterations: totalIterations,
        iterationsThisRun: localIteration,
        totalIterations: totalIterations,
        tolerance: tolerance,
        elapsedMs: Date.now() - startedAt,
        stagnationCount: stagnationCount,
        iterationProfile: iterationProfile
      }, comparison);
    }

    if (stagnationCount >= FS_FINANCE_ITERATION.stagnationLimit) {
      return Object.assign({
        converged: false,
        iterations: totalIterations,
        iterationsThisRun: localIteration,
        totalIterations: totalIterations,
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
    iterations: totalIterations,
    iterationsThisRun: iterationProfile.length,
    totalIterations: totalIterations,
    tolerance: tolerance,
    stoppedBySegmentLimit: true,
    elapsedMs: Date.now() - startedAt,
    stagnationCount: stagnationCount,
    iterationProfile: iterationProfile
  }, comparison);
}

/**
 * Giữ tên hàm cũ để các hàm khác trong dự án không bị lỗi tham chiếu.
 * Hàm này chạy theo phân đoạn dựa trên trạng thái hiện tại.
 */
function FS_hoiTuLaiVay_(runStartedAt) {
  return FS_hoiTuLaiVay_PhanDoan_(runStartedAt, FS_getRunState_());
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
      : result.reachedMaxIterations
        ? 'Đã đạt giới hạn tổng số vòng lặp.'
        : 'Phân đoạn hội tụ đã kết thúc; có thể chạy tiếp Bước 2.';

  return [
    'Mô hình tài trợ chưa hội tụ sau tổng cộng ' +
      (result.totalIterations || result.iterations || 0) + ' vòng.',
    reason,
    'Thời gian phiên hiện tại: ' + FS_formatSeconds_(result.elapsedMs || 0) + '.',
    'Sai lệch lớn nhất: ' + fmt(result.maxDelta) +
      ' đồng; ngưỡng: ' + fmt(result.tolerance) + ' đồng.',
    'Giải ngân: ' + fmt(result.drawDelta) + '; ' +
      'Lãi vay: ' + fmt(result.interestDelta) + '; ' +
      'Trả gốc: ' + fmt(result.principalDelta) + '; ' +
      'Dư nợ: ' + fmt(result.debtDelta) + '; ' +
      'Tiền cuối kỳ: ' + fmt(result.cashDelta) + ' đồng.'
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
  lines.push('Số vòng phiên này: ' + (result.iterationsThisRun || 0));
  lines.push('Tổng số vòng: ' + (result.totalIterations || result.iterations || 0));
  lines.push('Thời gian phiên: ' + FS_formatSeconds_(result.elapsedMs || 0));
  lines.push('Tổng thời gian hàm: ' + FS_formatSeconds_(Date.now() - runStartedAt));
  lines.push(
    'Sai lệch cuối: ' +
    Math.round(Number(result.maxDelta) || 0).toLocaleString('vi-VN') +
    ' đồng'
  );
  return lines.join('\n');
}

function FS_formatSeconds_(milliseconds) {
  return ((Number(milliseconds) || 0) / 1000).toFixed(2) + ' giây';
}

function FS_V22_KiemTraNhanhSauKhiDan() {
  const state = FS_getRunState_();
  const msg = [
    'Đã nạp runner FS V2.2 phân đoạn.',
    'Bước 1: khởi tạo Sheet 01, 03, 02 và 04.',
    'Bước 2: mỗi lần chạy tối đa ' +
      FS_FINANCE_ITERATION.maxIterationsPerRun +
      ' vòng hội tụ và lưu trạng thái.',
    'Bước 3: kiểm tra TNDN, lập Sheet 04A và Sheet 00.',
    'Trạng thái hiện tại: ' + (state.status || 'CHƯA KHỞI TẠO') + '.',
    'Tổng số vòng đã chạy: ' + state.totalIterations + '.',
    'Sai số tuyệt đối tối thiểu: ' +
      FS_FINANCE_ITERATION.absoluteTolerance.toLocaleString('vi-VN') +
      ' đồng.'
  ].join('\n');

  SpreadsheetApp.getUi().alert(msg);
}
