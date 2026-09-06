(function (root, factory) {
  "use strict";
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  root.TPTImportEngine = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  const text = (value) => String(value ?? "").trim();
  const keyText = (value) =>
    text(value)
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/đ/g, "d")
      .replace(/Đ/g, "D")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "");
  const compactText = (value) => text(value).replace(/\s+/g, " ");
  const codeText = (value) => compactText(value).toUpperCase();
  const numberValue = (value) => {
    if (value === "" || value == null) return null;
    const normalized = String(value).replace(/\s/g, "").replace(",", ".");
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : NaN;
  };
  const dateValue = (value) => {
    const raw = text(value);
    if (!raw) return "";
    if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
    const match = raw.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})$/);
    if (!match) return null;
    const day = Number(match[1]);
    const month = Number(match[2]);
    const year = Number(match[3]);
    const candidate = new Date(Date.UTC(year, month - 1, day));
    if (
      candidate.getUTCFullYear() !== year ||
      candidate.getUTCMonth() !== month - 1 ||
      candidate.getUTCDate() !== day
    )
      return null;
    return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  };

  const TYPES = {
    campuses: {
      label: "Cơ sở",
      store: "campuses",
      fields: [
        ["code", "Mã cơ sở", true, ["macoso", "ma_co_so", "campus_code"]],
        ["name", "Tên cơ sở", true, ["tencoso", "ten_co_so", "campus_name"]],
        ["address", "Địa chỉ", false, ["diachi", "dia_chi"]],
        ["status", "Trạng thái", false, ["trangthai", "trang_thai"]],
        ["notes", "Ghi chú", false, ["ghichu", "ghi_chu"]],
      ],
      key: (row) => codeText(row.code),
    },
    classes: {
      label: "Lớp và giáo viên chủ nhiệm",
      store: "classes",
      fields: [
        ["code", "Mã lớp", true, ["malop", "ma_lop", "class_code"]],
        ["class_name", "Tên lớp", true, ["tenlop", "ten_lop", "lop", "class_name"]],
        ["grade", "Khối", true, ["khoi", "grade"]],
        ["campus_code", "Mã cơ sở", true, ["macoso", "ma_co_so", "campus_code"]],
        ["campus_name", "Tên cơ sở", false, ["tencoso", "ten_co_so", "campus_name"]],
        ["teacher", "Giáo viên chủ nhiệm", false, ["giaovienchunhiem", "giao_vien_chu_nhiem", "gvcn", "teacher"]],
        ["size", "Sĩ số", false, ["siso", "si_so", "size"]],
        ["status", "Trạng thái", false, ["trangthai", "trang_thai"]],
        ["notes", "Ghi chú", false, ["ghichu", "ghi_chu"]],
      ],
      key: (row, context) =>
        [context.schoolYearId || row.school_year_id, codeText(row.campus_code), codeText(row.code || row.class_name)].join("|"),
    },
    homeroom_teachers: {
      label: "Giáo viên chủ nhiệm",
      store: "homeroom_teachers",
      fields: [
        ["code", "Mã giáo viên", false, ["magiaovien", "ma_giao_vien", "teacher_code"]],
        ["name", "Họ và tên", true, ["hovaten", "ho_va_ten", "tengiaovien", "name"]],
        ["class_code", "Mã lớp", true, ["malop", "ma_lop", "class_code"]],
        ["campus_code", "Mã cơ sở", true, ["macoso", "ma_co_so", "campus_code"]],
        ["phone", "Điện thoại", false, ["dienthoai", "dien_thoai", "phone"]],
        ["email", "Email", false, ["email"]],
        ["status", "Trạng thái", false, ["trangthai", "trang_thai"]],
      ],
      key: (row, context) =>
        [context.schoolYearId, codeText(row.campus_code), codeText(row.class_code)].join("|"),
    },
    students: {
      label: "Học sinh",
      store: "students",
      fields: [
        ["student_code", "Mã học sinh", true, ["mahocsinh", "ma_hoc_sinh", "student_code"]],
        ["full_name", "Họ và tên", true, ["hovaten", "ho_va_ten", "hoten", "student_name", "full_name"]],
        ["birth_date", "Ngày sinh", false, ["ngaysinh", "ngay_sinh", "birth_date"]],
        ["gender", "Giới tính", false, ["gioitinh", "gioi_tinh", "gender"]],
        ["class_code", "Mã lớp", true, ["malop", "ma_lop", "class_code"]],
        ["class_name", "Tên lớp", false, ["tenlop", "ten_lop", "lop", "class_name"]],
        ["campus_code", "Mã cơ sở", true, ["macoso", "ma_co_so", "campus_code"]],
        ["school_year", "Năm học", false, ["namhoc", "nam_hoc", "school_year"]],
        ["ordinal", "Số thứ tự", false, ["sothutu", "so_thu_tu", "stt", "ordinal"]],
        ["status", "Trạng thái", false, ["trangthai", "trang_thai", "status"]],
        ["organization_status", "Tình trạng Đội", false, ["tinhtrangdoi", "tinh_trang_doi", "tochucdoi", "organization_status"]],
        ["organization_joined_date", "Ngày kết nạp", false, ["ngayketnap", "ngay_ket_nap", "organization_joined_date"]],
        ["policy_groups", "Diện chính sách", false, ["dienchinhsach", "dien_chinh_sach", "chinhsach", "policy_groups"]],
        ["difficulty_status", "Tình trạng khó khăn", false, ["tinhtrangkhoKhan", "tinh_trang_kho_khan", "khokhan", "difficulty_status"]],
        ["special_needs", "Đặc điểm cần lưu ý", false, ["hocsinhdacbiet", "hoc_sinh_dac_biet", "dacdiemcanluuy", "special_needs"]],
        ["support_notes", "Hỗ trợ và theo dõi", false, ["hotrotheodoi", "ho_tro_theo_doi", "support_notes"]],
        ["guardian_name", "Người liên hệ", false, ["nguoilienhe", "nguoi_lien_he", "guardian_name"]],
        ["guardian_phone", "Điện thoại liên hệ", false, ["dienthoailienhe", "dien_thoai_lien_he", "guardian_phone"]],
        ["notes", "Ghi chú", false, ["ghichu", "ghi_chu", "notes"]],
      ],
      key: (row, context) =>
        [context.schoolYearId || row.school_year_id, codeText(row.student_code)].join("|"),
    },
    plans: {
      label: "Kế hoạch",
      store: "plans",
      fields: [
        ["code", "Mã kế hoạch", true, ["makehoach", "ma_ke_hoach", "code"]],
        ["name", "Tên kế hoạch", true, ["tenkehoach", "ten_ke_hoach", "name"]],
        ["level", "Cấp kế hoạch", false, ["capkehoach", "cap_ke_hoach", "level"]],
        ["start_date", "Ngày bắt đầu", false, ["ngaybatdau", "ngay_bat_dau", "start_date"]],
        ["end_date", "Ngày kết thúc", false, ["ngayketthuc", "ngay_ket_thuc", "end_date"]],
        ["objectives", "Mục tiêu", false, ["muctieu", "muc_tieu", "objectives"]],
        ["status", "Trạng thái", false, ["trangthai", "trang_thai", "status"]],
        ["progress", "Tiến độ (%)", false, ["tiendo", "tien_do", "progress"]],
        ["result", "Kết quả thực hiện", false, ["ketqua", "ket_qua", "result"]],
        ["evidence", "Minh chứng", false, ["minhchung", "minh_chung", "evidence"]],
        ["notes", "Ghi chú", false, ["ghichu", "ghi_chu", "notes"]],
      ],
      key: (row, context) => [context.schoolYearId, codeText(row.code)].join("|"),
    },
    tasks: {
      label: "Công việc và checklist",
      store: "tasks",
      fields: [
        ["code", "Mã công việc", false, ["macongviec", "ma_cong_viec", "code"]],
        ["title", "Tên công việc", true, ["tencongviec", "ten_cong_viec", "tieude", "title"]],
        ["description", "Mô tả", false, ["mota", "mo_ta", "description"]],
        ["plan_code", "Mã kế hoạch", false, ["makehoach", "ma_ke_hoach", "plan_code"]],
        ["start_date", "Ngày bắt đầu", false, ["ngaybatdau", "ngay_bat_dau", "start_date"]],
        ["due_date", "Hạn hoàn thành", false, ["hanhoanthanh", "han_hoan_thanh", "due_date"]],
        ["priority", "Ưu tiên", false, ["uutien", "uu_tien", "priority"]],
        ["assignee", "Người phụ trách", false, ["nguoiphutrach", "nguoi_phu_trach", "assignee"]],
        ["campus_code", "Mã cơ sở", false, ["macoso", "ma_co_so", "campus_code"]],
        ["status", "Trạng thái", false, ["trangthai", "trang_thai", "status"]],
        ["progress", "Tiến độ (%)", false, ["tiendo", "tien_do", "progress"]],
        ["completion_date", "Ngày hoàn thành", false, ["ngayhoanthanh", "ngay_hoan_thanh", "completion_date"]],
        ["result", "Kết quả thực hiện", false, ["ketqua", "ket_qua", "result"]],
        ["evidence", "Minh chứng", false, ["minhchung", "minh_chung", "evidence"]],
        ["notes", "Ghi chú", false, ["ghichu", "ghi_chu", "notes"]],
      ],
      key: (row, context) => [context.schoolYearId, codeText(row.code || row.title)].join("|"),
    },
    task_check_items: {
      label: "Checklist công việc",
      store: "task_check_items",
      fields: [
        ["task_code", "Mã công việc", true, ["macongviec", "ma_cong_viec", "task_code"]],
        ["label", "Nội dung checklist", true, ["noidung", "noi_dung", "tenmuc", "label"]],
        ["required", "Bắt buộc", false, ["batbuoc", "bat_buoc", "required"]],
        ["done", "Đã hoàn thành", false, ["dahoanthanh", "da_hoan_thanh", "done"]],
        ["order", "Thứ tự", false, ["thutu", "thu_tu", "order"]],
      ],
      key: (row, context) =>
        [context.schoolYearId, codeText(row.task_code), keyText(row.label)].join("|"),
    },
    calendar_events: {
      label: "Lịch hoạt động",
      store: "calendar_events",
      fields: [
        ["code", "Mã sự kiện", false, ["masukien", "ma_su_kien", "code"]],
        ["title", "Tên sự kiện", true, ["tensukien", "ten_su_kien", "tieude", "title"]],
        ["date", "Ngày", true, ["ngay", "date"]],
        ["start_time", "Giờ bắt đầu", false, ["giobatdau", "gio_bat_dau", "start_time"]],
        ["end_time", "Giờ kết thúc", false, ["gioketthuc", "gio_ket_thuc", "end_time"]],
        ["type", "Loại", false, ["loai", "type"]],
        ["location", "Địa điểm", false, ["diadiem", "dia_diem", "location"]],
        ["status", "Trạng thái", false, ["trangthai", "trang_thai", "status"]],
      ],
      key: (row, context) => [context.schoolYearId, row.date, row.start_time, keyText(row.title)].join("|"),
    },
    activities: {
      label: "Hoạt động Đội",
      store: "activities",
      fields: [
        ["code", "Mã hoạt động", false, ["mahoatdong", "ma_hoat_dong", "code"]],
        ["name", "Tên hoạt động", true, ["tenhoatdong", "ten_hoat_dong", "name"]],
        ["category", "Nhóm hoạt động", false, ["nhomhoatdong", "nhom_hoat_dong", "category"]],
        ["date", "Ngày tổ chức", true, ["ngaytochuc", "ngay_to_chuc", "date"]],
        ["location", "Địa điểm", false, ["diadiem", "dia_diem", "location"]],
        ["leader", "Người phụ trách", false, ["nguoiphutrach", "nguoi_phu_trach", "leader"]],
        ["participants", "Đối tượng/quy mô", false, ["doituong", "doi_tuong", "participants"]],
        ["status", "Trạng thái", false, ["trangthai", "trang_thai", "status"]],
      ],
      key: (row, context) => [context.schoolYearId, codeText(row.code || row.name), row.date].join("|"),
    },
    criteria: {
      label: "Tiêu chí thi đua",
      store: "criteria",
      fields: [
        ["code", "Mã tiêu chí", true, ["matieuchi", "ma_tieu_chi", "code"]],
        ["name", "Tên tiêu chí", true, ["tentieuchi", "ten_tieu_chi", "name"]],
        ["group", "Nhóm", false, ["nhom", "group"]],
        ["max_score", "Điểm tối đa", false, ["diemtoida", "diem_toi_da", "max_score"]],
        ["active", "Đang sử dụng", false, ["dangsudung", "dang_su_dung", "active"]],
      ],
      key: (row, context) => [context.schoolYearId, codeText(row.code)].join("|"),
    },
    score_entries: {
      label: "Điểm thi đua",
      store: "score_entries",
      fields: [
        ["class_code", "Mã lớp", true, ["malop", "ma_lop", "class_code"]],
        ["criterion_code", "Mã tiêu chí", true, ["matieuchi", "ma_tieu_chi", "criterion_code"]],
        ["week", "Tuần", true, ["tuan", "week"]],
        ["score", "Điểm", true, ["diem", "score"]],
        ["notes", "Ghi chú", false, ["ghichu", "ghi_chu", "notes"]],
      ],
      key: (row, context) => [context.schoolYearId, row.week, codeText(row.class_code), codeText(row.criterion_code)].join("|"),
    },
    student_incidents: {
      label: "Vi phạm học sinh",
      store: "student_incidents",
      fields: [
        ["incident_code", "Mã vi phạm", true, ["mavipham", "ma_vi_pham", "incident_code"]],
        ["date", "Ngày", true, ["ngay", "date"]],
        ["week", "Tuần", false, ["tuan", "week"]],
        ["student_code", "Mã học sinh", true, ["mahocsinh", "ma_hoc_sinh", "student_code"]],
        ["student_name", "Họ tên", false, ["hoten", "ho_ten", "hovaten", "student_name"]],
        ["class_code", "Mã lớp", true, ["malop", "ma_lop", "class_code"]],
        ["campus_code", "Mã cơ sở", true, ["macoso", "ma_co_so", "campus_code"]],
        ["incident_type", "Loại vi phạm", true, ["loaivipham", "loai_vi_pham", "incident_type"]],
        ["content", "Nội dung", true, ["noidung", "noi_dung", "content"]],
        ["proposed_deduction", "Điểm trừ đề xuất", false, ["diemtrudexuat", "diem_tru_de_xuat", "proposed_deduction"]],
        ["evidence", "Minh chứng", false, ["minhchung", "minh_chung", "evidence"]],
        ["recorded_by", "Người ghi nhận", false, ["nguoighinhan", "nguoi_ghi_nhan", "recorded_by"]],
        ["status", "Trạng thái", false, ["trangthai", "trang_thai", "status"]],
      ],
      key: (row) => codeText(row.incident_code),
    },
    team_members: {
      label: "Tổ chức Liên đội",
      store: "team_members",
      fields: [
        ["internal_code", "Mã", false, ["ma", "internal_code"]],
        ["name", "Họ và tên", true, ["hovaten", "ho_va_ten", "name"]],
        ["class_code", "Mã lớp", false, ["malop", "ma_lop", "class_code"]],
        ["unit", "Đội/ban", true, ["doiban", "doi_ban", "unit"]],
        ["position", "Chức vụ", true, ["chucvu", "chuc_vu", "position"]],
        ["term", "Nhiệm kỳ", false, ["nhiemky", "nhiem_ky", "term"]],
      ],
      key: (row, context) => [context.schoolYearId, codeText(row.internal_code || row.name), codeText(row.unit)].join("|"),
    },
    program_results: {
      label: "Rèn luyện – phong trào",
      store: "program_results",
      fields: [
        ["code", "Mã", false, ["ma", "code"]],
        ["name", "Tên chương trình", true, ["tenchuongtrinh", "ten_chuong_trinh", "name"]],
        ["scope", "Đối tượng/lớp", true, ["doituong", "doi_tuong", "scope"]],
        ["result", "Kết quả", false, ["ketqua", "ket_qua", "result"]],
        ["recognized_date", "Ngày công nhận", false, ["ngaycongnhan", "ngay_cong_nhan", "recognized_date"]],
        ["status", "Trạng thái", false, ["trangthai", "trang_thai", "status"]],
      ],
      key: (row, context) => [context.schoolYearId, codeText(row.code || row.name), keyText(row.scope)].join("|"),
    },
    commendations: {
      label: "Khen thưởng",
      store: "commendations",
      fields: [
        ["code", "Mã", false, ["ma", "code"]],
        ["award_type", "Loại khen thưởng", true, ["loaikhenthuong", "loai_khen_thuong", "award_type"]],
        ["recipient", "Đối tượng", true, ["doituong", "doi_tuong", "recipient"]],
        ["achievement", "Thành tích", true, ["thanhtich", "thanh_tich", "achievement"]],
        ["date", "Ngày", false, ["ngay", "date"]],
        ["approval_status", "Trạng thái", false, ["trangthai", "trang_thai", "approval_status"]],
      ],
      key: (row, context) => [context.schoolYearId, codeText(row.code || row.recipient), row.date].join("|"),
    },
    documents: {
      label: "Hồ sơ – minh chứng",
      store: "documents",
      fields: [
        ["code", "Mã hồ sơ", false, ["mahoso", "ma_ho_so", "code"]],
        ["name", "Tên hồ sơ", true, ["tenhoso", "ten_ho_so", "name"]],
        ["type", "Loại", false, ["loai", "type"]],
        ["date", "Ngày", false, ["ngay", "date"]],
        ["issuer", "Đơn vị ban hành", false, ["donvibanhanh", "don_vi_ban_hanh", "issuer"]],
        ["description", "Mô tả", false, ["mota", "mo_ta", "description"]],
        ["status", "Trạng thái", false, ["trangthai", "trang_thai", "status"]],
      ],
      key: (row, context) => [context.schoolYearId, codeText(row.code || row.name), row.date].join("|"),
    },
    equipment: {
      label: "Thiết bị",
      store: "equipment",
      fields: [
        ["code", "Mã thiết bị", true, ["mathietbi", "ma_thiet_bi", "code"]],
        ["name", "Tên thiết bị", true, ["tenthietbi", "ten_thiet_bi", "name"]],
        ["group", "Nhóm", false, ["nhom", "group"]],
        ["quantity", "Số lượng", true, ["soluong", "so_luong", "quantity"]],
        ["unit", "Đơn vị tính", false, ["donvitinh", "don_vi_tinh", "unit"]],
        ["condition", "Tình trạng", false, ["tinhtrang", "tinh_trang", "condition"]],
        ["location", "Nơi lưu", false, ["noiluu", "noi_luu", "location"]],
        ["notes", "Ghi chú", false, ["ghichu", "ghi_chu", "notes"]],
      ],
      key: (row) => codeText(row.code),
    },
    equipment_transactions: {
      label: "Giao nhận thiết bị",
      store: "equipment_transactions",
      fields: [
        ["code", "Mã giao nhận", true, ["magiaonhan", "ma_giao_nhan", "code"]],
        ["equipment_code", "Mã thiết bị", true, ["mathietbi", "ma_thiet_bi", "equipment_code"]],
        ["date", "Ngày", true, ["ngay", "date"]],
        ["type", "Loại giao nhận", true, ["loai", "type"]],
        ["quantity", "Số lượng", true, ["soluong", "so_luong", "quantity"]],
        ["receiver", "Người nhận", false, ["nguoinhan", "nguoi_nhan", "receiver"]],
        ["notes", "Ghi chú", false, ["ghichu", "ghi_chu", "notes"]],
      ],
      key: (row) => codeText(row.code),
    },
    generated_reports: {
      label: "Dữ liệu báo cáo",
      store: "generated_reports",
      fields: [
        ["code", "Mã báo cáo", false, ["mabaocao", "ma_bao_cao", "code"]],
        ["title", "Tên báo cáo", true, ["tenbaocao", "ten_bao_cao", "title"]],
        ["report_type", "Loại báo cáo", true, ["loaibaocao", "loai_bao_cao", "report_type"]],
        ["period", "Phạm vi", false, ["phamvi", "pham_vi", "period"]],
        ["status", "Trạng thái", false, ["trangthai", "trang_thai", "status"]],
        ["notes", "Ghi chú", false, ["ghichu", "ghi_chu", "notes"]],
      ],
      key: (row, context) => [context.schoolYearId, codeText(row.code || row.title), keyText(row.period)].join("|"),
    },
  };

  function parseDelimited(input, delimiter) {
    const source = String(input ?? "").replace(/^\uFEFF/, "");
    const selected = delimiter || (source.includes("\t") ? "\t" : ",");
    const rows = [];
    let row = [];
    let cell = "";
    let quoted = false;
    for (let index = 0; index < source.length; index += 1) {
      const char = source[index];
      if (char === '"') {
        if (quoted && source[index + 1] === '"') {
          cell += '"';
          index += 1;
        } else quoted = !quoted;
      } else if (char === selected && !quoted) {
        row.push(cell);
        cell = "";
      } else if ((char === "\n" || char === "\r") && !quoted) {
        if (char === "\r" && source[index + 1] === "\n") index += 1;
        row.push(cell);
        if (row.some((value) => text(value))) rows.push(row);
        row = [];
        cell = "";
      } else cell += char;
    }
    row.push(cell);
    if (row.some((value) => text(value))) rows.push(row);
    return rows;
  }

  function detectHeaderRow(rows, type, requestedIndex) {
    if (Number.isInteger(requestedIndex) && requestedIndex >= 0) return requestedIndex;
    const schema = TYPES[type];
    if (!schema) return 0;
    const aliases = new Set(
      schema.fields.flatMap(([name, label, , alternatives]) => [
        keyText(name),
        keyText(label),
        ...(alternatives || []).map(keyText),
      ]),
    );
    let best = { index: 0, score: -1 };
    rows.slice(0, 10).forEach((row, index) => {
      const score = row.reduce(
        (total, value) => total + (aliases.has(keyText(value)) ? 1 : 0),
        0,
      );
      if (score > best.score) best = { index, score };
    });
    return best.index;
  }

  function autoMap(headers, type) {
    const schema = TYPES[type];
    if (!schema) return {};
    const normalized = headers.map(keyText);
    const mapping = {};
    for (const [name, label, , alternatives] of schema.fields) {
      const aliases = [name, label, ...(alternatives || [])].map(keyText);
      const index = normalized.findIndex((header) => aliases.includes(header));
      if (index >= 0) mapping[name] = index;
    }
    return mapping;
  }

  function mapRows(rows, type, headerIndex, mapping, defaults = {}) {
    const schema = TYPES[type];
    if (!schema) throw new Error("Loại dữ liệu không được hỗ trợ.");
    return rows.slice(headerIndex + 1).map((row, offset) => {
      const mapped = { ...defaults };
      for (const [name] of schema.fields) {
        if (Number.isInteger(mapping[name])) mapped[name] = compactText(row[mapping[name]]);
        else if (mapped[name] == null) mapped[name] = "";
      }
      return { sourceRow: headerIndex + offset + 2, values: mapped, source: row };
    });
  }

  function normalizeMappedRow(type, row, context = {}) {
    const output = { ...row };
    for (const key of Object.keys(output)) {
      if (typeof output[key] === "string") output[key] = compactText(output[key]);
    }
    for (const name of ["code", "student_code", "incident_code", "class_code", "campus_code", "criterion_code", "equipment_code"])
      if (name in output) output[name] = codeText(output[name]);
    for (const name of ["birth_date", "start_date", "end_date", "due_date", "completion_date", "date", "recognized_date", "organization_joined_date"])
      if (name in output && output[name]) output[name] = dateValue(output[name]);
    for (const name of ["grade", "size", "ordinal", "max_score", "score", "proposed_deduction", "quantity", "order", "progress"])
      if (name in output && output[name] !== "") output[name] = numberValue(output[name]);
    if (type === "classes" || type === "students" || type === "homeroom_teachers")
      output.school_year_id = context.schoolYearId || output.school_year_id || "";
    if (type === "students") {
      const organizationKey = keyText(output.organization_status);
      output.organization_status =
        ({
          nhi_dong: "nhi_dong",
          du_bi_doi_vien: "du_bi",
          du_bi: "du_bi",
          doi_vien: "doi_vien",
          doan_vien: "doan_vien",
          chua_xac_dinh: "chua_xac_dinh",
          "": "",
        })[organizationKey] || output.organization_status;
      const difficultyKey = keyText(output.difficulty_status);
      output.difficulty_status =
        ({
          khong: "none",
          khong_co: "none",
          can_theo_doi: "monitor",
          can_ho_tro: "support",
          can_ho_tro_khan: "urgent",
          "": "",
        })[difficultyKey] || output.difficulty_status;
    }
    if (type !== "campuses" && context.schoolYearId) {
      output.school_year_id ||= context.schoolYearId;
      output.academic_year_id ||= context.schoolYearId;
    }
    return output;
  }

  function validateRows(type, mappedRows, options = {}) {
    const schema = TYPES[type];
    if (!schema) throw new Error("Loại dữ liệu không được hỗ trợ.");
    const context = options.context || {};
    const campusCodes = new Set((options.campuses || []).map((row) => codeText(row.code)));
    const classCodes = new Set((options.classes || []).map((row) => codeText(row.code || row.class_name)));
    const studentCodes = new Set((options.students || []).map((row) => codeText(row.student_code)));
    const existingKeys = new Set(
      (options.existing || []).map((row) => schema.key(row, context)).filter(Boolean),
    );
    const seen = new Set();
    const seenNames = new Set();
    const existingStudentNames = new Set(
      type === "students"
        ? (options.existing || []).map((row) =>
            [codeText(row.class_code || row.class_name), keyText(row.full_name)].join("|"),
          )
        : [],
    );
    const existingClassNames = new Set(
      type === "classes"
        ? (options.existing || []).map((row) =>
            [codeText(row.campus_code || row.campus_id), keyText(row.class_name)].join("|"),
          )
        : [],
    );
    return mappedRows.map((item) => {
      const row = normalizeMappedRow(type, item.values, context);
      const errors = [];
      const warnings = [];
      const blankFields = [];
      for (const [name, label, required] of schema.fields) {
        if (row[name] === "" || row[name] == null)
          blankFields.push({ name, label, required: Boolean(required) });
      }
      if (blankFields.length)
        warnings.push(
          `Đang để trống: ${blankFields.map((field) => field.label).join(", ")}`,
        );
      for (const name of ["birth_date", "start_date", "end_date", "due_date", "completion_date", "date", "recognized_date", "organization_joined_date"])
        if (name in row && row[name] === null) errors.push(`Sai định dạng ngày ở ${name}`);
      for (const name of ["grade", "size", "ordinal", "max_score", "score", "proposed_deduction", "quantity", "order", "progress"])
        if (name in row && Number.isNaN(row[name])) errors.push(`Sai kiểu số ở ${name}`);
      if (type === "classes" && row.grade != null && !Number.isNaN(row.grade) && (row.grade < 1 || row.grade > 9))
        errors.push("Khối phải từ 1 đến 9");
      if (row.campus_code && campusCodes.size && !campusCodes.has(codeText(row.campus_code)))
        errors.push("Không tìm thấy mã cơ sở");
      if (["students", "student_incidents", "homeroom_teachers", "score_entries"].includes(type) && row.class_code && classCodes.size && !classCodes.has(codeText(row.class_code)))
        errors.push("Không tìm thấy lớp");
      if (type === "student_incidents" && row.student_code && studentCodes.size && !studentCodes.has(codeText(row.student_code)))
        warnings.push("Không tìm thấy mã học sinh; cần kiểm tra trước khi nhập");
      if (type === "equipment" && Number.isFinite(row.quantity) && row.quantity < 0)
        errors.push("Số lượng không được âm");
      if (type === "students") {
        if (
          row.organization_status &&
          ![
            "chua_xac_dinh",
            "nhi_dong",
            "du_bi",
            "doi_vien",
            "doan_vien",
          ].includes(row.organization_status)
        )
          errors.push("Tình trạng Đội không hợp lệ");
        if (
          row.difficulty_status &&
          !["none", "monitor", "support", "urgent"].includes(
            row.difficulty_status,
          )
        )
          errors.push("Tình trạng khó khăn không hợp lệ");
        if (row.full_name) {
          const nameKey = [codeText(row.class_code), keyText(row.full_name)].join("|");
          if (seenNames.has(nameKey)) warnings.push("Trùng họ tên trong cùng lớp trong tệp");
          if (existingStudentNames.has(nameKey)) warnings.push("Đã có học sinh trùng họ tên trong lớp; đối chiếu mã học sinh");
          seenNames.add(nameKey);
        }
      }
      if (type === "classes" && row.class_name) {
        const nameKey = [codeText(row.campus_code), keyText(row.class_name)].join("|");
        if (seenNames.has(nameKey) || existingClassNames.has(nameKey))
          errors.push("Trùng tên lớp trong cùng cơ sở và năm học");
        seenNames.add(nameKey);
      }
      const duplicateKey = schema.key(row, context);
      const canCheckDuplicate = !blankFields.some((field) => field.required);
      const duplicateInFile = canCheckDuplicate && duplicateKey && seen.has(duplicateKey);
      const exists = canCheckDuplicate && duplicateKey && existingKeys.has(duplicateKey);
      if (duplicateInFile) errors.push("Trùng khóa trong tệp");
      if (canCheckDuplicate && duplicateKey) seen.add(duplicateKey);
      if (exists) warnings.push("Bản ghi đã tồn tại và có thể cập nhật");
      return {
        ...item,
        values: row,
        duplicateKey,
        exists: Boolean(exists),
        valid: errors.length === 0,
        errors,
        warnings,
        blankFields,
      };
    });
  }

  async function checksum(input) {
    const bytes =
      input instanceof ArrayBuffer
        ? input
        : ArrayBuffer.isView(input)
          ? input.buffer.slice(input.byteOffset, input.byteOffset + input.byteLength)
          : new TextEncoder().encode(String(input ?? "")).buffer;
    if (globalThis.crypto?.subtle) {
      const digest = await globalThis.crypto.subtle.digest("SHA-256", bytes);
      return [...new Uint8Array(digest)].map((value) => value.toString(16).padStart(2, "0")).join("");
    }
    let hash = 2166136261;
    for (const value of new Uint8Array(bytes)) hash = Math.imul(hash ^ value, 16777619);
    return `fnv1a-${(hash >>> 0).toString(16).padStart(8, "0")}`;
  }

  function csv(rows, columns) {
    const safe = (value) => {
      let output = String(value ?? "");
      if (/^[=+\-@]/.test(output)) output = `'${output}`;
      return `"${output.replace(/"/g, '""')}"`;
    };
    const body = [columns.map((column) => safe(column.label)).join(",")];
    for (const row of rows) body.push(columns.map((column) => safe(row[column.name])).join(","));
    return `\uFEFF${body.join("\r\n")}`;
  }

  return Object.freeze({
    TYPES,
    text,
    keyText,
    compactText,
    codeText,
    dateValue,
    numberValue,
    parseDelimited,
    detectHeaderRow,
    autoMap,
    mapRows,
    normalizeMappedRow,
    validateRows,
    checksum,
    csv,
  });
});
