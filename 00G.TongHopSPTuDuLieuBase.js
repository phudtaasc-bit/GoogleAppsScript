function FS_SP_tongHopTrucTiepTuBase() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName('01A. Kỹ thuật');
  const summary = ss.getSheetByName('00. Tổng hợp');

  if (!tech || !summary) {
    throw new Error('Thiếu Sheet 01A. Kỹ thuật hoặc Sheet 00. Tổng hợp.');
  }

  const products = typeof FS00P_docSanPham_ === 'function'
    ? FS00P_docSanPham_(tech)
    : FS_SP_BASE_docSanPham_(tech);

  if (!products.length) {
    throw new Error('Block SAN_PHAM không có sản phẩm hợp lệ.');
  }

  const message = products
    .map(product => product.code + ' - ' + product.name + (product.group ? ' - ' + product.group : ''))
    .join('\n');

  const response = ui.prompt(
    'Tổng hợp hiệu quả từng sản phẩm',
    'Nhập Mã SP cần kiểm tra. Có thể nhập riêng mã hoặc dán cả dòng:\n\n' + message,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const raw = String(response.getResponseText() || '').trim();
  const code = raw.split('-')[0].trim().toUpperCase();
  const product = products.find(item => item.code === code);

  if (!product) {
    throw new Error('Mã SP "' + raw.toUpperCase() + '" không tồn tại trong block SAN_PHAM.');
  }

  // Hàm này chỉ đọc dữ liệu Base đã hội tụ từ 02, 03 và 03A.
  // Không dựng kịch bản, không ghi 0 vào sheet nguồn và không chạy hội tụ 03A ↔ 04.
  if (typeof FS00P_capNhatCotSanPham_ !== 'function') {
    throw new Error('Không tìm thấy hàm tổng hợp trực tiếp FS00P_capNhatCotSanPham_.');
  }

  FS00P_capNhatCotSanPham_();
  SpreadsheetApp.flush();

  const headerRow = 24;
  const label = product.name + (product.group ? ' - ' + product.group : '');
  const headers = summary.getRange(headerRow, 1, 1, summary.getLastColumn()).getDisplayValues()[0];
  const productCol = headers.findIndex(value => FS_SP_BASE_key_(value) === FS_SP_BASE_key_(label)) + 1;

  if (productCol > 0) {
    summary.activate();
    summary.setActiveSelection(summary.getRange(headerRow, productCol));
  }

  ui.alert(
    'Đã tổng hợp ' + product.code + ' - ' + label + ' từ dữ liệu Base.\n\n' +
    'Lãi vay lấy nguyên giá trị "Lãi vay vốn hóa phân bổ" tại Sheet 03A.\n' +
    'Không tính lại lãi vay, không hội tụ Sheet 04 và không thay đổi chỉ tiêu toàn dự án.'
  );
}

function FS_SP_BASE_docSanPham_(sheet) {
  if (typeof FS00_findBlock_ !== 'function') {
    throw new Error('Không tìm thấy hàm FS00_findBlock_ để đọc block SAN_PHAM.');
  }

  const block = FS00_findBlock_(sheet, 'SAN_PHAM');
  const values = sheet.getRange(block.startRow, 1, block.rowCount, 14).getValues();
  const products = [];
  const seen = {};

  values.forEach(row => {
    const code = String(row[0] || '').trim().toUpperCase();
    const name = String(row[1] || '').trim();
    const group = String(row[2] || '').trim();
    if (!code || !name || seen[code]) return;
    seen[code] = true;
    products.push({ code, name, group });
  });

  return products;
}

function FS_SP_BASE_key_(value) {
  return String(value == null ? '' : value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}
