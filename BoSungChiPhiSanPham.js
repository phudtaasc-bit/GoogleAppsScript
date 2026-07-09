/*************************************************
 * FS V2.1 - CLEAN FILE
 * Dán ĐÈ TOÀN BỘ nội dung file BoSungChiPhiSanPham... hiện tại bằng file này.
 *
 * Mục tiêu:
 * - Xóa tác dụng của patch V4/V4.1 cũ.
 * - Không ghi đè FS_lapSheet02, FS_lapSheet03, FS_lapSheet00.
 * - Chỉ giữ hàm chạy mô hình và hàm chạy tách bước.
 *
 * LƯU Ý:
 * - File 02.Doanhthu.gs phải dùng bản FS_V2_0_02_DoanhThu.
 * - File 03.Chiphi.gs phải dùng bản FS_V2_1_03_ChiPhiVonVay.
 *************************************************/

function FS_chayToanBoMoHinh_CoLaiVay() {
  const ss = SpreadsheetApp.getActive();
  ss.toast('Đang chạy toàn bộ mô hình...', 'FS V2.1', 5);

  FS_taoKyThuatTuDauVao();

  FS_lapSheet03();
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
  FS_lapSheet03();
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
 * Nếu kết quả báo OK nhưng vẫn lỗi, nghĩa là còn file khác đang chứa hàm cũ
 * và cần tìm bằng Ctrl + Shift + F.
 */
function FS_V21_KiemTraNhanhSauKhiDan() {
  const ss = SpreadsheetApp.getActive();
  const msg = [
    'Đã nạp file CLEAN V2.1.',
    'File này không còn ghi đè FS_lapSheet02 / FS_lapSheet03 / FS_lapSheet00.',
    'Hãy bảo đảm:',
    '1) 02.Doanhthu.gs đã dán bản V2.0.',
    '2) 03.Chiphi.gs đã dán bản V2.1.',
    '3) File BoSungChiPhiSanPham... chỉ còn nội dung CLEAN này.'
  ].join('\n');

  SpreadsheetApp.getUi().alert(msg);
}
