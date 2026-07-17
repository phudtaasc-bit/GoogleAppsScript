function onOpen() {
  SpreadsheetApp.getUi().createMenu('FS - MÔ HÌNH ĐẦU TƯ')
    .addItem('1. Cập nhật 01A. Kỹ thuật', 'FS_capNhatKyThuat')
    .addItem('2. Lập 02. Doanh thu', 'FS_lapSheet02')
    .addItem('3. Lập 03. Chi phí & Vốn', 'FS_lapSheet03')
    .addItem('4. Lập 03A. Lợi nhuận & Thuế', 'FS_lapSheet03A')
    .addItem('5. Hội tụ 04. Dòng tiền & Tài trợ', 'FS_hoiTuTaiTro')
    .addSeparator()
    .addItem('6. Lập 04A. Tổng hợp dòng tiền', 'FS_lapSheet04A')
    .addItem('7. Lập 00. Tổng hợp', 'FS_lapSheet00_TheoDanhMuc')
    .addItem('8. Lập bảng độ nhạy', 'FS05_lapBangDoNhay_AnToan')
    .addItem('9. Chạy 99. Kiểm tra mô hình', 'FS_lapSheet99')
    .addSeparator()
    .addItem('Chạy toàn bộ mô hình', 'FS_chayToanBo')
    .addSeparator()
    .addItem('Phân tích hiệu quả từng sản phẩm', 'FS00H_phanTichVaDinhDangTheoSanPham')
    .addItem('Tạo bảng kiểm tra IRR sản phẩm', 'FS00G_taoBangDebugIRR')
    .addToUi();
}

function FS_chayToanBo() {
  FS_capNhatKyThuat();
  FS_lapSheet02();
  FS_lapSheet03();
  FS_hoiTuTaiTro();
  if (typeof FS_lapSheet04A === 'function') FS_lapSheet04A();
  if (typeof FS_lapSheet00_TheoDanhMuc === 'function') FS_lapSheet00_TheoDanhMuc();
  FS_lapSheet99();
  SpreadsheetApp.getUi().alert('Đã chạy xong mô hình và kiểm tra.');
}
