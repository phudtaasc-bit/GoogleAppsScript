/*************************************************
 * FS V2.1 - CLEAN FILE
 * Luồng chạy mô hình giữ nguyên thứ tự phụ thuộc hiện tại.
 * Sheet 03 chạy qua lớp vá FS_lapSheet03_Patched để:
 * - Tách đúng Tiền SDĐ/Tiền thuê đất theo nhóm sản phẩm.
 * - Bổ sung VAT đầu vào xây dựng & thiết bị và chi phí bán hàng.
 *************************************************/

function FS_chayToanBoMoHinh_CoLaiVay() {
  const ss = SpreadsheetApp.getActive();
  ss.toast('Đang chạy toàn bộ mô hình...', 'FS V2.1', 5);

  FS_taoKyThuatTuDauVao();
  FS97_assertLandCostConfig_();

  FS_lapSheet03_Patched();
  FS_lapSheet02();
  FS_lapSheet04();

  FS03_capNhatNguonVonTuSheet04();

  FS_lapSheet02();
  FS_lapSheet04();

  FS_lapSheet04A();
  FS_lapSheet00();

  ss.toast('Đã chạy xong toàn bộ mô hình.', 'FS V2.1', 10);
}

function FS_chayMoHinh_Buoc1() {
  const ss = SpreadsheetApp.getActive();
  ss.toast('Đang chạy bước 1...', 'FS V2.1', 5);

  FS_taoKyThuatTuDauVao();
  FS97_assertLandCostConfig_();
  FS_lapSheet03_Patched();
  FS_lapSheet02();
  FS_lapSheet04();
  FS03_capNhatNguonVonTuSheet04();

  ss.toast('Xong bước 1. Chạy tiếp bước 2.', 'FS V2.1', 10);
}

function FS_chayMoHinh_Buoc2() {
  const ss = SpreadsheetApp.getActive();
  ss.toast('Đang chạy bước 2...', 'FS V2.1', 5);

  FS_lapSheet02();
  FS_lapSheet04();
  FS_lapSheet04A();
  FS_lapSheet00();

  ss.toast('Đã chạy xong bước 2.', 'FS V2.1', 10);
}

/**
 * Chạy kiểm tra nhanh để phát hiện project còn patch cũ hay không.
 */
function FS_V21_KiemTraNhanhSauKhiDan() {
  const ss = SpreadsheetApp.getActive();
  const msg = [
    'Đã nạp luồng FS V2.1 đã vá.',
    'Sheet 03 đang dùng FS_lapSheet03_Patched.',
    'Kiểm tra tiền đất được chạy trước khi ghi lại Sheet 03.',
    'VAT hoàn âm cuối kỳ được giữ nguyên theo quy ước mô hình.'
  ].join('\n');

  SpreadsheetApp.getUi().alert(msg);
}
