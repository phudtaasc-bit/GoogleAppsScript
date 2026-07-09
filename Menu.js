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
    .addItem('4.3. Chạy toàn bộ mô hình - có cập nhật lãi vay', 'FS_chayToanBoMoHinh_CoLaiVay')
    .addSeparator()
    .addItem('5. Độ nhạy nhanh', 'FS05_DoNhay_Fast')
    .addItem('6. Độ nhạy vốn CSH - NPV/IRR vốn', 'FS05B_DoNhay_CSH')

    .addToUi();
}
function FS05B_DoNhay_CSH() {
  boSungDoNhay_NPV_IRR_VonCSH();
}
function FS_chayToanBo() {
  FS_taoKyThuatTuDauVao();
  FS_lapSheet02();
  FS_lapSheet03();
  FS_lapSheet04();
  FS_lapSheet00();
  SpreadsheetApp.getUi().alert('Đã cập nhật toàn bộ mô hình.');
}
function FS_chayMoHinh_Buoc1() {
  const ss = SpreadsheetApp.getActive();
  ss.toast('Đang chạy bước 1...', 'FS', 5);

  FS_taoKyThuatTuDauVao();
  FS_lapSheet03();
  FS_lapSheet02();
  FS_lapSheet04();
  FS03_capNhatNguonVonTuSheet04();

  ss.toast('Xong bước 1. Chạy tiếp 4.3B.', 'FS', 10);
}

function FS_chayMoHinh_Buoc2() {
  const ss = SpreadsheetApp.getActive();
  ss.toast('Đang chạy bước 2...', 'FS', 5);

  FS_lapSheet02();
  FS_lapSheet04();
  FS_lapSheet04A();
  FS_lapSheet00();

  ss.toast('Đã chạy xong toàn bộ mô hình.', 'FS', 10);
}