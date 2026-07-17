const FSSP_STEP_PROPERTY = 'FS_SP_STEP_V2';

function FS_SP_B1_ChuanBiKichBan() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(3000)) throw new Error('Mô hình đang được một tiến trình khác sử dụng.');

  try {
    const ss = SpreadsheetApp.getActive();
    const props = PropertiesService.getDocumentProperties();
    if (props.getProperty(FSSP_CFG.BACKUP_PROPERTY)) {
      throw new Error('Đang tồn tại bản sao của lần chạy trước. Hãy chạy "0. Khôi phục mô hình gốc" trước.');
    }

    const tech = ss.getSheetByName(FSSP_CFG.TECH);
    if (!tech) throw new Error('Không tìm thấy Sheet 01A. Kỹ thuật.');
    const products = FS_SP_docSanPham_(tech);
    if (!products.length) throw new Error('Block SAN_PHAM không có sản phẩm hợp lệ.');

    const ui = SpreadsheetApp.getUi();
    const list = products.map(p => p.code + ' - ' + p.name + (p.group ? ' - ' + p.group : '')).join('\n');
    const response = ui.prompt(
      'Bước 1 - Chuẩn bị kịch bản sản phẩm',
      'Nhập Mã SP hoặc dán cả dòng trong danh sách:\n\n' + list,
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() !== ui.Button.OK) return;

    const raw = String(response.getResponseText() || '').trim();
    const code = raw.split(' - ')[0].trim().toUpperCase();
    const product = products.find(p => p.code === code);
    if (!product) throw new Error('Mã SP "' + code + '" không tồn tại trong block SAN_PHAM.');

    const sourceNames = [
      FSSP_CFG.REVENUE,
      FSSP_CFG.COST,
      FSSP_CFG.PROFIT,
      FSSP_CFG.CASH,
      FSSP_CFG.CASH_SUMMARY,
      FSSP_CFG.SUMMARY
    ].filter(name => ss.getSheetByName(name));

    const backups = FS_SP_taoBanSao_(ss, sourceNames);
    props.setProperty(FSSP_CFG.BACKUP_PROPERTY, JSON.stringify(backups));
    props.setProperty(FSSP_STEP_PROPERTY, JSON.stringify({
      code: product.code,
      name: product.name,
      group: product.group,
      stage: 'PREPARED'
    }));

    FS_SP_apDungKichBan_(ss.getSheetByName(FSSP_CFG.REVENUE), product.code, 'REVENUE');
    FS_SP_apDungKichBan_(ss.getSheetByName(FSSP_CFG.COST), product.code, 'COST');
    SpreadsheetApp.flush();

    ui.alert(
      'Đã chuẩn bị kịch bản ' + product.code + ' - ' + product.name + '.\n' +
      'Tiếp theo chạy: 2. Lập 03A cho sản phẩm.'
    );
  } catch (error) {
    FS_SP_Buoc_TuKhoiPhucNeuCan_();
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function FS_SP_B2_Lap03A() {
  FS_SP_Buoc_Chay_('PREPARED', 'PROFIT_DONE', function() {
    if (typeof FS_lapSheet03A !== 'function') throw new Error('Không tìm thấy hàm FS_lapSheet03A.');
    FS_lapSheet03A();
  }, 'Đã lập 03A. Tiếp theo chạy: 3. Hội tụ 04 cho sản phẩm.');
}

function FS_SP_B3_HoiTu04() {
  FS_SP_Buoc_Chay_('PROFIT_DONE', 'CASH_DONE', function() {
    if (typeof FS_hoiTuTaiTro !== 'function') throw new Error('Không tìm thấy hàm FS_hoiTuTaiTro.');
    FS_hoiTuTaiTro();
  }, 'Đã hội tụ 04. Tiếp theo chạy: 4. Ghi kết quả và khôi phục Base.');
}

function FS_SP_B4_GhiKetQuaVaKhoiPhuc() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(3000)) throw new Error('Mô hình đang được một tiến trình khác sử dụng.');

  const ss = SpreadsheetApp.getActive();
  const props = PropertiesService.getDocumentProperties();
  try {
    const state = FS_SP_Buoc_DocTrangThai_();
    if (state.stage !== 'CASH_DONE') {
      throw new Error('Sai trình tự. Trạng thái hiện tại: ' + state.stage + '.');
    }

    if (typeof FS_lapSheet04A === 'function') FS_lapSheet04A();
    if (typeof FS_lapSheet00_TheoDanhMuc === 'function') {
      FS_lapSheet00_TheoDanhMuc();
    } else if (typeof FS_lapSheet00 === 'function') {
      FS_lapSheet00();
    } else {
      throw new Error('Không tìm thấy hàm lập Sheet 00.');
    }
    SpreadsheetApp.flush();

    const scenarioSummary = ss.getSheetByName(FSSP_CFG.SUMMARY);
    const scenarioRows = FS_SP_docDongChiTieu_(scenarioSummary);
    const scenarioValues = scenarioSummary
      .getRange(scenarioRows.totalRevenue, 4, scenarioRows.vatPayable - scenarioRows.totalRevenue + 1, 1)
      .getValues();

    const backupJson = props.getProperty(FSSP_CFG.BACKUP_PROPERTY);
    if (!backupJson) throw new Error('Không tìm thấy bản sao để khôi phục mô hình Base.');
    FS_SP_khoiPhucTuBanSao_(ss, JSON.parse(backupJson));
    props.deleteProperty(FSSP_CFG.BACKUP_PROPERTY);
    props.deleteProperty(FSSP_STEP_PROPERTY);
    props.deleteProperty(FSSP_CFG.STATUS_PROPERTY);

    const tech = ss.getSheetByName(FSSP_CFG.TECH);
    const products = FS_SP_docSanPham_(tech);
    FS_SP_chuanHoaCotSanPhamV2_(ss, products);

    const baseSummary = ss.getSheetByName(FSSP_CFG.SUMMARY);
    const baseRows = FS_SP_docDongChiTieu_(baseSummary);
    const productIndex = products.findIndex(p => p.code === state.code);
    if (productIndex < 0) throw new Error('Không xác định được cột của sản phẩm ' + state.code + '.');

    const targetCol = 5 + productIndex;
    baseSummary
      .getRange(baseRows.totalRevenue, targetCol, scenarioValues.length, 1)
      .setValues(scenarioValues);
    FS_SP_dinhDangCotKetQua_(baseSummary, baseRows, targetCol);
    SpreadsheetApp.flush();

    SpreadsheetApp.getUi().alert(
      'Đã ghi kết quả riêng cho ' + state.code + ' - ' + state.name + '.\n' +
      'Mô hình tổng dự án đã được khôi phục.'
    );
  } catch (error) {
    FS_SP_Buoc_TuKhoiPhucNeuCan_();
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function FS_SP_Buoc_Chay_(requiredStage, nextStage, action, successMessage) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(3000)) throw new Error('Mô hình đang được một tiến trình khác sử dụng.');

  try {
    const props = PropertiesService.getDocumentProperties();
    const state = FS_SP_Buoc_DocTrangThai_();
    if (state.stage !== requiredStage) {
      throw new Error('Sai trình tự. Cần trạng thái ' + requiredStage + ', hiện tại là ' + state.stage + '.');
    }

    action();
    SpreadsheetApp.flush();
    state.stage = nextStage;
    props.setProperty(FSSP_STEP_PROPERTY, JSON.stringify(state));
    SpreadsheetApp.getUi().alert(successMessage);
  } finally {
    lock.releaseLock();
  }
}

function FS_SP_Buoc_DocTrangThai_() {
  const raw = PropertiesService.getDocumentProperties().getProperty(FSSP_STEP_PROPERTY);
  if (!raw) throw new Error('Chưa chuẩn bị kịch bản sản phẩm. Hãy chạy bước 1 trước.');
  try {
    return JSON.parse(raw);
  } catch (error) {
    throw new Error('Trạng thái kịch bản không hợp lệ. Hãy khôi phục mô hình gốc.');
  }
}

function FS_SP_Buoc_TuKhoiPhucNeuCan_() {
  const ss = SpreadsheetApp.getActive();
  const props = PropertiesService.getDocumentProperties();
  const backupJson = props.getProperty(FSSP_CFG.BACKUP_PROPERTY);
  if (!backupJson) return;

  try {
    FS_SP_khoiPhucTuBanSao_(ss, JSON.parse(backupJson));
    props.deleteProperty(FSSP_CFG.BACKUP_PROPERTY);
    props.deleteProperty(FSSP_STEP_PROPERTY);
    props.deleteProperty(FSSP_CFG.STATUS_PROPERTY);
    SpreadsheetApp.flush();
  } catch (restoreError) {
    // Giữ nguyên bản sao để người dùng có thể chạy menu khôi phục thủ công.
  }
}

function FS_SP_khoiPhucMoHinhGoc_V2() {
  FS_SP_khoiPhucMoHinhGoc();
  const props = PropertiesService.getDocumentProperties();
  props.deleteProperty(FSSP_STEP_PROPERTY);
  props.deleteProperty(FSSP_CFG.STATUS_PROPERTY);
}
