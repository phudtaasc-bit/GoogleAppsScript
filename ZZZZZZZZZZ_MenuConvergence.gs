/*************************************************
 * ZZZZZZZZZZ_MenuConvergence.gs
 * Chuẩn hóa menu vận hành mô hình:
 * - Bỏ các bước hội tụ thủ công 4.3.1 -> 4.3.5.
 * - Chỉ giữ một lệnh chạy toàn bộ mô hình đến hội tụ.
 *************************************************/

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
    .addItem('4.3 Chạy toàn bộ mô hình - Hội tụ lãi vay & thuế', 'FS_chayToanBoMoHinh_CoLaiVay')

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
