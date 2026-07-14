/*************************************************
 * 99R_AuditAnnualLabelFix.js
 * Vá chính xác việc đọc nhãn động tại Sheet 04A.
 *
 * Sheet 04A có thể dùng nhãn mô tả dài, ví dụ:
 * "Dòng tiền thuần sau tài trợ = Tổng dòng tiền vào - Tổng dòng tiền ra".
 * Audit chỉ yêu cầu phần tên chỉ tiêu ở đầu nhãn.
 *************************************************/

function FS99Q_annualValue_(annual, label) {
  const target = FS99Q_key_(label);

  // 1. Ưu tiên khớp tuyệt đối.
  if (Object.prototype.hasOwnProperty.call(annual, target)) {
    return FS99Q_num_(annual[target]);
  }

  // 2. Chỉ chấp nhận nhãn thực tế bắt đầu bằng tên chỉ tiêu cần tìm.
  // Không dùng contains hai chiều để tránh lấy nhầm dòng có nội dung gần giống.
  const matches = Object.keys(annual).filter(key =>
    key === target || key.indexOf(target + ' ') === 0
  );

  if (matches.length === 1) {
    return FS99Q_num_(annual[matches[0]]);
  }

  if (matches.length > 1) {
    throw new Error(
      'Audit 04A: Nhãn "' + label + '" khớp nhiều dòng: ' + matches.join(' | ')
    );
  }

  throw new Error(
    'Audit 04A: Không tìm thấy dòng chỉ tiêu "' + label + '" tại cột Nội dung.'
  );
}
