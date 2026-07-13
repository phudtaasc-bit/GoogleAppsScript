function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('FS - CẬP NHẬT MÔ HÌNH')

    .addItem('0. Lập sheet 00 - Tổng hợp', 'FS_lapSheet00')
    .addItem('1. Tạo lại sheet Kỹ thuật từ Đầu vào', 'FS_taoKyThuatTuDauVao')

    .addSeparator()
    .addItem('2. Lập sheet 02 - Doanh thu', 'FS_lapSheet02')

    .addSeparator()
    .addItem('3. Lập sheet 03 - Chi phí & vốn vay', 'FS_lapSheet03_Safe')
    .addItem('3.1 Sheet 03 - Chế độ quản trị', 'FS03_viewQuanTri')
    .addItem('3.2 Sheet 03 - Hiển thị đầy đủ', 'FS03_viewDayDu')
    .addItem('3.3 Sheet 03 - Định dạng V2', 'FS03_formatV2')
    .addItem('3.4 Sheet 03 - Kiểm tra logic', 'FS03_checkV2')

    .addSeparator()
    .addItem('4. Lập sheet 04 - Dòng tiền', 'FS_lapSheet04')
    .addItem('4.1 Sheet 04 - Kiểm tra logic', 'FS04_check')
    .addItem('4.2 Sheet 04A - Tổng hợp dòng tiền', 'FS_lapSheet04A')

    .addSeparator()
    .addItem('4.3.1 Bước 1 - Khởi tạo mô hình', 'FS_chayMoHinh_Buoc1')
    .addItem('4.3.2 Bước 2 - Tiếp tục hội tụ tài trợ', 'FS_chayMoHinh_Buoc2')
    .addItem('4.3.3 Bước 3 - Hoàn tất và lập báo cáo', 'FS_chayMoHinh_Buoc3')
    .addItem('4.3.4 Xóa trạng thái chạy', 'FS_xoaTrangThaiChay')
    .addItem('4.3.5 Chạy tự động theo trạng thái', 'FS_chayToanBoMoHinh_Safe')

    .addSeparator()
    .addItem('5. Độ nhạy nhanh', 'FS05_DoNhay_Fast')
    .addItem('6. Độ nhạy vốn CSH - NPV/IRR vốn', 'FS05B_DoNhay_CSH')
    .addItem('6.1 Độ nhạy vốn CSH - Lãi suất & Vốn đầu tư', 'FS05D_DoNhay_CSH_LaiSuat_VonDauTu')

    .addSeparator()
    .addItem('7. Kiểm thử hồi quy toàn mô hình', 'FS93_RunRegressionSuite')

    .addToUi();
}

/**
 * Alias tương thích cho các nút hoặc trigger cũ.
 * Luồng chạy chuẩn nằm tại BoSungChiPhiSanPham.js và có cập nhật lãi vay.
 */
function FS_chayToanBo() {
  return FS_chayToanBoMoHinh_Safe();
}
