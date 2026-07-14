/*************************************************
 * ZZZZZZZZZZZZ_StatefulConvergence.gs
 * Chia quy trình hội tụ thành các lần chạy ngắn để không vượt giới hạn Apps Script.
 *************************************************/

const FS_CONV_STATE_KEY_ = 'FS_CONV_STATE_V2';
const FS_CONV_MAX_ITER_ = 10;
const FS_CONV_TOL_ = 1000; // đồng

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('FS - CẬP NHẬT MÔ HÌNH')
    .addItem('0. Lập sheet 00 - Tổng hợp', 'FS_lapSheet00')
    .addItem('1. Tạo lại sheet Kỹ thuật từ Đầu vào', 'FS_taoKyThuatTuDauVao')
    .addSeparator()
    .addItem('2. Lập sheet 02 - Doanh thu', 'FS_lapSheet02')
    .addSeparator()
    .addItem('3. Lập sheet 03 - Chi phí & vốn vay', 'FS_lapSheet03')
    .addItem('3.1 Sheet 03 - Chế độ quản trị', 'FS03_viewQuanTri')
    .addItem('3.2 Sheet 03 - Hiển thị đầy đủ', 'FS03_viewDayDu')
    .addItem('3.3 Sheet 03 - Định dạng V2', 'FS03_formatV2')
    .addItem('3.4 Sheet 03 - Kiểm tra logic', 'FS03_checkV2')
    .addSeparator()
    .addItem('4. Lập sheet 04 - Dòng tiền', 'FS_lapSheet04')
    .addItem('4.1 Sheet 04 - Kiểm tra logic', 'FS04_check')
    .addItem('4.2 Sheet 04A - Tổng hợp dòng tiền', 'FS_lapSheet04A')
    .addSeparator()
    .addItem('4.3.1 Hội tụ - Khởi tạo', 'FS_CONV_khoiTao')
    .addItem('4.3.2 Hội tụ - Chạy vòng tiếp theo', 'FS_CONV_chayVongTiepTheo')
    .addItem('4.3.3 Hội tụ - Hoàn tất báo cáo', 'FS_CONV_hoanTatBaoCao')
    .addItem('4.3.4 Hội tụ - Xem trạng thái', 'FS_CONV_xemTrangThai')
    .addItem('4.3.5 Hội tụ - Xóa trạng thái', 'FS_CONV_xoaTrangThai')
    .addSeparator()
    .addItem('8.1 Pipeline - Cập nhật chi phí nguồn', 'FSZZZZZZZ_pipelineCapNhatChiPhiNguon')
    .addItem('8.2 Pipeline - Lập dòng tiền', 'FSZZZZZZZ_pipelineLapDongTien')
    .addItem('8.3 Pipeline - Tổng hợp theo năm', 'FSZZZZZZZ_pipelineTongHopNam')
    .addItem('8.4 Pipeline - Tổng hợp dự án', 'FSZZZZZZZ_pipelineTongHopDuAn')
    .addItem('8.5 Pipeline - Audit phân rã sai lệch', 'FS99Q_chayAuditReconciliation')
    .addSeparator()
    .addItem('5. Độ nhạy nhanh', 'FS05_DoNhay_Fast')
    .addItem('6. Độ nhạy vốn CSH - NPV/IRR vốn', 'FS05B_DoNhay_CSH')
    .addItem('6.1 Độ nhạy vốn CSH - Lãi suất & Vốn đầu tư', 'FS05B_DoNhay_CSH')
    .addToUi();
}

function FS_CONV_khoiTao() {
  FS_CONV_requireFunctions_();

  FS_taoKyThuatTuDauVao();
  FS_lapSheet03();
  FS_lapSheet02();
  SpreadsheetApp.flush();
  FS_lapSheet04();
  SpreadsheetApp.flush();
  FS_CONV_sync04To03_();

  const state = {
    iteration: 1,
    previousInterest: FS_CONV_sumNamed_('04. Dòng tiền & Lợi nhuận', ['Lãi vay vốn hóa', 'Lãi vay']),
    previousCit: FS_CONV_sumNamed_('02. Doanh thu', ['Thuế TNDN tạm tính']),
    currentInterest: null,
    currentCit: null,
    diffInterest: null,
    diffCit: null,
    converged: false,
    finalized: false,
    updatedAt: new Date().toISOString()
  };
  FS_CONV_saveState_(state);

  SpreadsheetApp.getUi().alert(
    'Đã khởi tạo vòng 1.\n' +
    'Tiếp theo chạy menu 4.3.2 cho đến khi trạng thái báo ĐÃ HỘI TỤ.'
  );
}

function FS_CONV_chayVongTiepTheo() {
  FS_CONV_requireFunctions_();
  const state = FS_CONV_loadState_();
  if (!state) throw new Error('Chưa có trạng thái hội tụ. Hãy chạy 4.3.1 trước.');
  if (state.finalized) throw new Error('Mô hình đã hoàn tất báo cáo. Muốn chạy lại, dùng 4.3.5 rồi chạy 4.3.1.');
  if (state.converged) {
    SpreadsheetApp.getUi().alert('Mô hình đã hội tụ. Chạy 4.3.3 để hoàn tất báo cáo.');
    return;
  }
  if (state.iteration >= FS_CONV_MAX_ITER_) {
    throw new Error('Đã đạt tối đa ' + FS_CONV_MAX_ITER_ + ' vòng nhưng chưa hội tụ.');
  }

  FS_lapSheet02();
  SpreadsheetApp.flush();
  FS_lapSheet04();
  SpreadsheetApp.flush();
  FS_CONV_sync04To03_();

  const interest = FS_CONV_sumNamed_('04. Dòng tiền & Lợi nhuận', ['Lãi vay vốn hóa', 'Lãi vay']);
  const cit = FS_CONV_sumNamed_('02. Doanh thu', ['Thuế TNDN tạm tính']);
  const diffInterest = Math.abs(interest - Number(state.previousInterest || 0));
  const diffCit = Math.abs(cit - Number(state.previousCit || 0));

  state.iteration += 1;
  state.currentInterest = interest;
  state.currentCit = cit;
  state.diffInterest = diffInterest;
  state.diffCit = diffCit;
  state.converged = diffInterest <= FS_CONV_TOL_ && diffCit <= FS_CONV_TOL_;
  state.previousInterest = interest;
  state.previousCit = cit;
  state.updatedAt = new Date().toISOString();
  FS_CONV_saveState_(state);

  SpreadsheetApp.getUi().alert(
    'Đã chạy vòng ' + state.iteration + '.\n' +
    'Chênh lệch lãi vay: ' + FS_CONV_format_(diffInterest) + ' đồng.\n' +
    'Chênh lệch Thuế TNDN: ' + FS_CONV_format_(diffCit) + ' đồng.\n' +
    (state.converged
      ? 'Trạng thái: ĐÃ HỘI TỤ. Chạy 4.3.3 để hoàn tất báo cáo.'
      : 'Trạng thái: CHƯA HỘI TỤ. Tiếp tục chạy 4.3.2.')
  );
}

function FS_CONV_hoanTatBaoCao() {
  const state = FS_CONV_loadState_();
  if (!state) throw new Error('Chưa có trạng thái hội tụ. Hãy chạy 4.3.1 trước.');
  if (!state.converged) {
    throw new Error(
      'Mô hình chưa hội tụ. Vòng hiện tại: ' + state.iteration +
      '; lệch lãi vay: ' + FS_CONV_format_(state.diffInterest) +
      '; lệch Thuế TNDN: ' + FS_CONV_format_(state.diffCit) + ' đồng.'
    );
  }

  // Lập lại lần cuối trên bộ lãi vay đã hội tụ.
  FS_lapSheet02();
  SpreadsheetApp.flush();
  FS_lapSheet04();
  SpreadsheetApp.flush();
  FS_CONV_sync04To03_();

  if (typeof FS_lapSheet04A === 'function') FS_lapSheet04A();
  if (typeof FS_lapSheet00 === 'function') FS_lapSheet00();
  if (typeof FS99Q_chayAuditReconciliation === 'function') FS99Q_chayAuditReconciliation();

  state.finalized = true;
  state.updatedAt = new Date().toISOString();
  FS_CONV_saveState_(state);

  SpreadsheetApp.getUi().alert(
    'Đã hoàn tất mô hình sau ' + state.iteration + ' vòng hội tụ.\n' +
    'Đã lập Sheet 04A, Sheet 00 và chạy Audit Sheet 99.'
  );
}

function FS_CONV_xemTrangThai() {
  const state = FS_CONV_loadState_();
  if (!state) {
    SpreadsheetApp.getUi().alert('Chưa có trạng thái hội tụ.');
    return;
  }
  SpreadsheetApp.getUi().alert(
    'Vòng hiện tại: ' + state.iteration + '/' + FS_CONV_MAX_ITER_ + '\n' +
    'Lãi vay: ' + FS_CONV_format_(state.previousInterest) + ' đồng\n' +
    'Thuế TNDN: ' + FS_CONV_format_(state.previousCit) + ' đồng\n' +
    'Lệch lãi vay: ' + FS_CONV_format_(state.diffInterest) + ' đồng\n' +
    'Lệch Thuế TNDN: ' + FS_CONV_format_(state.diffCit) + ' đồng\n' +
    'Hội tụ: ' + (state.converged ? 'CÓ' : 'CHƯA') + '\n' +
    'Hoàn tất báo cáo: ' + (state.finalized ? 'CÓ' : 'CHƯA')
  );
}

function FS_CONV_xoaTrangThai() {
  PropertiesService.getDocumentProperties().deleteProperty(FS_CONV_STATE_KEY_);
  SpreadsheetApp.getUi().alert('Đã xóa trạng thái hội tụ.');
}

// Giữ tương thích nếu nơi khác vẫn gọi hàm cũ: chỉ khởi tạo, không chạy dài gây timeout.
function FS_chayToanBoMoHinh_CoLaiVay() {
  FS_CONV_khoiTao();
}

function FS_CONV_sync04To03_() {
  if (typeof FS03_capNhatNguonVonTuSheet04 === 'function') {
    FS03_capNhatNguonVonTuSheet04();
    SpreadsheetApp.flush();
  }
}

function FS_CONV_saveState_(state) {
  PropertiesService.getDocumentProperties().setProperty(FS_CONV_STATE_KEY_, JSON.stringify(state));
}

function FS_CONV_loadState_() {
  const raw = PropertiesService.getDocumentProperties().getProperty(FS_CONV_STATE_KEY_);
  return raw ? JSON.parse(raw) : null;
}

function FS_CONV_requireFunctions_() {
  ['FS_taoKyThuatTuDauVao', 'FS_lapSheet03', 'FS_lapSheet02', 'FS_lapSheet04'].forEach(name => {
    if (typeof globalThis[name] !== 'function') throw new Error('Thiếu hàm ' + name + '.');
  });
}
