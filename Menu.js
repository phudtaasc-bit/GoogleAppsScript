function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const productMenu = ui.createMenu('Tổng hợp hiệu quả từng SP')
    .addItem('Chọn sản phẩm để tính', 'FS_SP_chonVaTinhMotSanPham_V2')
    .addItem('Khôi phục mô hình gốc', 'FS_SP_khoiPhucMoHinhGoc');

  ui.createMenu('FS - MÔ HÌNH ĐẦU TƯ')
    .addItem('1. Cập nhật 01A. Kỹ thuật', 'FS_capNhatKyThuat')
    .addItem('2. Lập 02. Doanh thu', 'FS_lapSheet02')
    .addItem('3. Lập 03. Chi phí & Vốn', 'FS_lapSheet03')
    .addItem('4. Lập 03A. Lợi nhuận & Thuế', 'FS_lapSheet03A')
    .addItem('5. Hội tụ 04. Dòng tiền & Tài trợ', 'FS_hoiTuTaiTro')
    .addSeparator()
    .addItem('6. Lập 04A. Tổng hợp dòng tiền', 'FS_lapSheet04A')
    .addItem('7. Lập 00. Tổng hợp', 'FS_lapSheet00_TheoDanhMucVaSanPham_GhiChuCuoi')
    .addItem('8. Lập bảng độ nhạy', 'FS05_lapBangDoNhay_AnToan')
    .addItem('9. Chạy 99. Kiểm tra mô hình', 'FS_lapSheet99')
    .addSeparator()
    .addSubMenu(productMenu)
    .addSeparator()
    .addItem('Chạy toàn bộ mô hình', 'FS_chayToanBo')
    .addToUi();
}

function FS_chayToanBo() {
  FS_capNhatKyThuat();
  FS_lapSheet02();
  FS_lapSheet03();
  FS_hoiTuTaiTro();
  if (typeof FS_lapSheet04A === 'function') FS_lapSheet04A();
  if (typeof FS_lapSheet00_TheoDanhMucVaSanPham_GhiChuCuoi === 'function') {
    FS_lapSheet00_TheoDanhMucVaSanPham_GhiChuCuoi();
  }
  FS_lapSheet99();
  SpreadsheetApp.getUi().alert('Đã chạy xong mô hình và kiểm tra.');
}
