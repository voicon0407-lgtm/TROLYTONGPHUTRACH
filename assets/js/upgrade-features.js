(function (root) {
  "use strict";

  const SCRIPT_PROMISES = new Map();

  function create(context) {
    const {
      APP,
      state,
      db,
      engine,
      $,
      $$,
      uid,
      now,
      today,
      esc,
      fmtDate,
      fmtDateTime,
      normalizeText,
      campusName,
      pageHead,
      setContent,
      toast,
      openModal,
      closeModal,
      download,
      loadContext,
      createInternalSnapshot,
      saveDocumentFiles,
      classForm,
      setting,
      go,
    } = context;

    if (!engine) throw new Error("Thiếu bộ xử lý nhập dữ liệu cục bộ.");

    state.studentPage ||= 1;
    state.studentPageSize ||= 50;
    state.studentQuery ||= "";
    state.studentClassFilter ||= "all";
    state.studentStatusFilter ||= "all";
    state.studentOrganizationFilter ||= "all";
    state.studentSupportFilter ||= "all";
    state.studentSelected ||= new Set();
    state.incidentPage ||= 1;
    state.incidentPageSize ||= 50;
    state.incidentQuery ||= "";
    state.incidentClassFilter ||= "all";
    state.incidentStatusFilter ||= "all";
    state.classPage ||= 1;
    state.classPageSize ||= 50;
    state.classQuery ||= "";
    state.classGradeFilter ||= "all";
    state.classStatusFilter ||= "all";
    state.classSelected ||= new Set();

    const organizationLabel = (value) =>
      ({
        nhi_dong: "Nhi đồng",
        du_bi: "Dự bị đội viên",
        doi_vien: "Đội viên",
        doan_vien: "Đoàn viên",
        chua_xac_dinh: "Chưa xác định",
      })[value] || value || "Chưa xác định";
    const difficultyLabel = (value) =>
      ({
        none: "Không",
        monitor: "Cần theo dõi",
        support: "Cần hỗ trợ",
        urgent: "Cần hỗ trợ khẩn",
      })[value] || value || "Không";

    const escapeRegExp = (value) =>
      String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const safeFileName = (value) =>
      String(value || "du-lieu")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/đ/g, "d")
        .replace(/Đ/g, "D")
        .replace(/[<>:"/\\|?*\x00-\x1F]/g, "-")
        .replace(/\s+/g, "-")
        .replace(/-+/g, "-")
        .replace(/^-|-$/g, "")
        .slice(0, 120);
    const currentYear = () =>
      state.cache.years?.find((row) => row.id === state.yearId) || null;
    const currentCampus = () =>
      state.cache.campuses?.find((row) => row.id === state.campusId) || null;
    const pageSlice = (rows, page, size) => {
      const pages = Math.max(1, Math.ceil(rows.length / size));
      const safePage = Math.min(pages, Math.max(1, Number(page) || 1));
      return {
        rows: rows.slice((safePage - 1) * size, safePage * size),
        page: safePage,
        pages,
        total: rows.length,
      };
    };
    const paginationHtml = (id, model) =>
      `<div class="pagination"><span class="muted">${model.total.toLocaleString("vi-VN")} bản ghi • Trang ${model.page}/${model.pages}</span><button class="btn small" id="${id}Prev" ${model.page <= 1 ? "disabled" : ""}>‹ Trước</button><button class="btn small" id="${id}Next" ${model.page >= model.pages ? "disabled" : ""}>Sau ›</button></div>`;

    function loadScriptOnce(relativePath, globalName) {
      if (globalName && root[globalName]) return Promise.resolve(root[globalName]);
      const src = new URL(relativePath, location.href).href;
      if (!SCRIPT_PROMISES.has(src)) {
        SCRIPT_PROMISES.set(
          src,
          new Promise((resolve, reject) => {
            const existing = [...document.scripts].find((script) => script.src === src);
            if (existing) {
              if (!globalName || root[globalName]) return resolve(root[globalName]);
              existing.addEventListener("load", () => resolve(root[globalName]), {
                once: true,
              });
              existing.addEventListener("error", () => reject(new Error(`Không tải được ${relativePath}`)), {
                once: true,
              });
              return;
            }
            const script = document.createElement("script");
            script.src = src;
            script.async = true;
            script.onload = () => resolve(globalName ? root[globalName] : true);
            script.onerror = () => reject(new Error(`Không tải được ${relativePath}`));
            document.head.append(script);
          }),
        );
      }
      return SCRIPT_PROMISES.get(src);
    }

    async function loadXlsx() {
      const lib = await loadScriptOnce("./assets/vendor/xlsx/xlsx.full.min.js", "XLSX");
      if (!lib) throw new Error("Thư viện Excel cục bộ chưa sẵn sàng.");
      return lib;
    }

    async function loadMammoth() {
      const lib = await loadScriptOnce(
        "./assets/vendor/mammoth/mammoth.browser.min.js",
        "mammoth",
      );
      if (!lib) throw new Error("Thư viện Word cục bộ chưa sẵn sàng.");
      return lib;
    }

    async function loadPdfJs() {
      if (root.__TPT_PDFJS__) return root.__TPT_PDFJS__;
      if (!Uint8Array.prototype.toHex) {
        Object.defineProperty(Uint8Array.prototype, "toHex", {
          configurable: true,
          value() {
            return Array.from(this, (byte) => byte.toString(16).padStart(2, "0")).join("");
          },
        });
      }
      const moduleUrl = new URL(
        "./assets/vendor/pdfjs/pdf.min.mjs",
        location.href,
      ).href;
      const pdfjs = await import(moduleUrl);
      pdfjs.GlobalWorkerOptions.workerSrc = new URL(
        "./assets/vendor/pdfjs/pdf.worker.min.mjs",
        location.href,
      ).href;
      root.__TPT_PDFJS__ = pdfjs;
      return pdfjs;
    }

    async function loadTesseract() {
      const lib = await loadScriptOnce(
        "./assets/vendor/tesseract/tesseract.min.js",
        "Tesseract",
      );
      if (!lib) throw new Error("Bộ OCR cục bộ chưa sẵn sàng.");
      return lib;
    }

    async function exportXlsx(rows, columns, fileName, title) {
      const XLSX = await loadXlsx();
      const data = rows.map((row) =>
        Object.fromEntries(columns.map((column) => [column.label, row[column.name] ?? ""])),
      );
      const sheet = XLSX.utils.json_to_sheet(data, { cellDates: true });
      sheet["!cols"] = columns.map((column) => ({
        wch: Math.min(
          42,
          Math.max(
            column.label.length + 2,
            ...rows.slice(0, 200).map((row) => String(row[column.name] ?? "").length + 1),
          ),
        ),
      }));
      if (rows.length)
        sheet["!autofilter"] = { ref: `A1:${XLSX.utils.encode_col(columns.length - 1)}${rows.length + 1}` };
      const overview = XLSX.utils.aoa_to_sheet([
        [title || "Dữ liệu xuất"],
        ["Trường", APP.schoolName],
        ["Năm học", currentYear()?.name || "Tất cả"],
        ["Cơ sở", currentCampus()?.name || "Tất cả"],
        ["Thời gian xuất", new Date().toLocaleString("vi-VN")],
        ["Số bản ghi", rows.length],
      ]);
      overview["!cols"] = [{ wch: 20 }, { wch: 48 }];
      const book = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(book, overview, "TỔNG QUAN");
      XLSX.utils.book_append_sheet(book, sheet, "DỮ LIỆU");
      const array = XLSX.write(book, { bookType: "xlsx", type: "array" });
      download(
        new Blob([array], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }),
        `${safeFileName(fileName)}.xlsx`,
      );
    }

    function exportCsv(rows, columns, fileName) {
      download(
        engine.csv(rows, columns),
        `${safeFileName(fileName)}.csv`,
        "text/csv;charset=utf-8",
      );
    }

    async function exportWordTable(rows, columns, fileName, title) {
      const table = `<table><thead><tr>${columns.map((column) => `<th>${esc(column.label)}</th>`).join("")}</tr></thead><tbody>${rows
        .map(
          (row) =>
            `<tr>${columns.map((column) => `<td>${esc(row[column.name] ?? "")}</td>`).join("")}</tr>`,
        )
        .join("")}</tbody></table>`;
      const documentHtml = `<!doctype html><html lang="vi"><head><meta charset="utf-8"><style>body{font-family:Arial,sans-serif;font-size:11pt}h1{text-align:center;font-size:16pt}table{width:100%;border-collapse:collapse}th,td{border:1px solid #555;padding:5px;vertical-align:top}th{background:#eaf4ff}.sign{margin-top:36px;text-align:right;padding-right:50px}</style></head><body><p><strong>${esc(APP.schoolName)}</strong></p><h1>${esc(title)}</h1><p>Năm học: ${esc(currentYear()?.name || "Tất cả")} • Cơ sở: ${esc(currentCampus()?.name || "Tất cả")}</p>${table}<div class="sign"><strong>NGƯỜI LẬP BÁO CÁO</strong><br><br><br><br>(Ký và ghi rõ họ tên)</div></body></html>`;
      download(
        `\uFEFF${documentHtml}`,
        `${safeFileName(fileName)}.doc`,
        "application/msword;charset=utf-8",
      );
    }

    async function migrateSmartImportData() {
      if (await setting("migration_10_completed")) return;
      const operation = await db.put(
        "operation_journal",
        {
          operation_id: uid(),
          operation: "schema_migration_v10",
          status: "started",
          from_schema: Number(db.upgradedFrom ?? 9),
          to_schema: APP.schema,
          started_at: now(),
        },
        { audit: false, journal: false },
      );
      try {
        const before = await createInternalSnapshot("Trước hoàn thiện schema 10", {
          tier: "protected",
          protectedSnapshot: true,
          reason: "before-migration-v10",
        });
        await db.put(
          "migration_logs",
          {
            from_schema: Number(db.upgradedFrom ?? 9),
            to_schema: 10,
            status: "success",
            snapshot_id: before.id,
            summary:
              "Bổ sung học sinh, vi phạm và trung tâm nhập dữ liệu; không đổi ID hoặc record cũ.",
          },
          { audit: false },
        );
        await setting("migration_10_completed", true);
        await createInternalSnapshot("Sau hoàn thiện schema 10", {
          tier: "protected",
          protectedSnapshot: true,
          reason: "after-migration-v10",
        });
        await db.put(
          "operation_journal",
          { ...operation, status: "completed", finished_at: now() },
          { audit: false, journal: false },
        );
      } catch (error) {
        await db
          .put(
            "migration_logs",
            {
              from_schema: Number(db.upgradedFrom ?? 9),
              to_schema: 10,
              status: "failed",
              summary: error.message,
            },
            { audit: false },
          )
          .catch(() => {});
        await db
          .put(
            "operation_journal",
            {
              ...operation,
              status: "failed",
              error: error.message,
              finished_at: now(),
            },
            { audit: false, journal: false },
          )
          .catch(() => {});
        throw error;
      }
    }

    async function studentsForView() {
      const [students, classes] = await Promise.all([
        db.all("students"),
        db.all("classes"),
      ]);
      const classById = new Map(classes.map((row) => [row.id, row]));
      const query = normalizeText(state.studentQuery);
      return students
        .filter((row) => !state.yearId || row.school_year_id === state.yearId)
        .filter((row) => state.campusId === "all" || row.campus_id === state.campusId)
        .filter(
          (row) =>
            state.studentClassFilter === "all" ||
            row.class_id === state.studentClassFilter,
        )
        .filter(
          (row) =>
            state.studentStatusFilter === "all" ||
            (row.status || "active") === state.studentStatusFilter,
        )
        .filter(
          (row) =>
            state.studentOrganizationFilter === "all" ||
            (row.organization_status || "chua_xac_dinh") ===
              state.studentOrganizationFilter,
        )
        .filter((row) => {
          if (state.studentSupportFilter === "all") return true;
          if (state.studentSupportFilter === "policy")
            return !!String(row.policy_groups || "").trim();
          if (state.studentSupportFilter === "difficulty")
            return !["", "none"].includes(row.difficulty_status || "none");
          if (state.studentSupportFilter === "special")
            return !!String(row.special_needs || "").trim();
          return true;
        })
        .filter((row) => {
          if (!query) return true;
          const klass = classById.get(row.class_id);
          return normalizeText(
            [
              row.student_code,
              row.full_name,
              row.class_code,
              klass?.class_name,
              row.organization_status,
              row.policy_groups,
              row.difficulty_status,
              row.special_needs,
              row.support_notes,
              row.notes,
            ]
              .filter(Boolean)
              .join(" "),
          ).includes(query);
        })
        .sort((a, b) =>
          `${a.class_code || ""}|${a.ordinal || 9999}|${a.full_name || ""}`.localeCompare(
            `${b.class_code || ""}|${b.ordinal || 9999}|${b.full_name || ""}`,
            "vi",
            { numeric: true },
          ),
        );
    }

    async function renderStudents() {
      const [rows, classes] = await Promise.all([
        studentsForView(),
        db.all("classes"),
      ]);
      const availableClasses = classes
        .filter((row) => !state.yearId || row.school_year_id === state.yearId)
        .filter((row) => state.campusId === "all" || row.campus_id === state.campusId)
        .sort((a, b) => String(a.class_name).localeCompare(String(b.class_name), "vi", { numeric: true }));
      const model = pageSlice(rows, state.studentPage, state.studentPageSize);
      state.studentPage = model.page;
      const selected = state.studentSelected;
      setContent(
        pageHead(
          "Học sinh",
          "Quản lý theo năm học, cơ sở và lớp; danh sách được phân trang để dùng tốt với trên 4.500 học sinh.",
          `<button class="btn" id="studentTemplate">Tải mẫu Excel</button><button class="btn" id="studentImport">Nhập dữ liệu</button><button class="btn primary" id="studentAdd">＋ Học sinh</button>`,
        ) +
          `<div class="metric-row"><div class="metric"><strong>${rows.length.toLocaleString("vi-VN")}</strong><span>Học sinh theo bộ lọc</span></div><div class="metric"><strong>${rows.filter((row) => row.organization_status === "nhi_dong").length}</strong><span>Nhi đồng</span></div><div class="metric"><strong>${rows.filter((row) => row.organization_status === "doi_vien").length}</strong><span>Đội viên</span></div><div class="metric"><strong>${rows.filter((row) => row.organization_status === "doan_vien").length}</strong><span>Đoàn viên</span></div><div class="metric"><strong>${rows.filter((row) => String(row.policy_groups || "").trim()).length}</strong><span>Diện chính sách</span></div><div class="metric"><strong>${rows.filter((row) => !["", "none"].includes(row.difficulty_status || "none") || String(row.special_needs || "").trim()).length}</strong><span>Cần theo dõi/hỗ trợ</span></div></div>
          <div class="card mt"><div class="card-body">
          <div class="toolbar"><input class="grow" id="studentSearch" value="${esc(state.studentQuery)}" placeholder="Tìm mã, tên, lớp, chính sách hoặc ghi chú…"><select id="studentClassFilter"><option value="all">Tất cả lớp</option>${availableClasses.map((row) => `<option value="${row.id}" ${state.studentClassFilter === row.id ? "selected" : ""}>${esc(row.class_name)}</option>`).join("")}</select><select id="studentStatusFilter"><option value="all">Tất cả trạng thái học</option><option value="active" ${state.studentStatusFilter === "active" ? "selected" : ""}>Đang học</option><option value="inactive" ${state.studentStatusFilter === "inactive" ? "selected" : ""}>Ngừng học</option><option value="transferred" ${state.studentStatusFilter === "transferred" ? "selected" : ""}>Chuyển trường</option><option value="graduated" ${state.studentStatusFilter === "graduated" ? "selected" : ""}>Hoàn thành cấp học</option></select><select id="studentOrganizationFilter"><option value="all">Mọi tình trạng Đội</option>${[["nhi_dong","Nhi đồng"],["du_bi","Dự bị đội viên"],["doi_vien","Đội viên"],["doan_vien","Đoàn viên"],["chua_xac_dinh","Chưa xác định"]].map(([value,label]) => `<option value="${value}" ${state.studentOrganizationFilter === value ? "selected" : ""}>${label}</option>`).join("")}</select><select id="studentSupportFilter"><option value="all">Mọi diện theo dõi</option><option value="policy" ${state.studentSupportFilter === "policy" ? "selected" : ""}>Diện chính sách</option><option value="difficulty" ${state.studentSupportFilter === "difficulty" ? "selected" : ""}>Có hoàn cảnh khó khăn</option><option value="special" ${state.studentSupportFilter === "special" ? "selected" : ""}>Có đặc điểm cần lưu ý</option></select><select id="studentPageSize"><option value="25" ${state.studentPageSize === 25 ? "selected" : ""}>25 dòng</option><option value="50" ${state.studentPageSize === 50 ? "selected" : ""}>50 dòng</option><option value="100" ${state.studentPageSize === 100 ? "selected" : ""}>100 dòng</option></select><button class="btn small" id="studentCsv">CSV</button><button class="btn small" id="studentXlsx">Excel</button><button class="btn small" id="studentWord">Word</button></div>
          ${selected.size ? `<div class="bulkbar"><strong>${selected.size} học sinh đã chọn</strong><button class="btn small" id="studentBulkTransfer">Chuyển lớp</button><button class="btn small" id="studentBulkInactive">Ngừng học</button><button class="btn small" id="studentBulkRestore">Khôi phục</button><button class="btn small" id="studentClearSelection">Bỏ chọn</button></div>` : ""}
          <div class="table-wrap" style="max-height:520px"><table><thead><tr><th><input type="checkbox" id="studentSelectPage" aria-label="Chọn trang hiện tại"></th><th>STT</th><th>Mã học sinh</th><th>Họ và tên</th><th>Lớp</th><th>Tình trạng Đội</th><th>Chính sách/hỗ trợ</th><th>Trạng thái học</th><th>Thao tác</th></tr></thead><tbody>${model.rows
            .map(
              (row, index) =>
                `<tr><td><input type="checkbox" data-student-select="${row.id}" ${selected.has(row.id) ? "checked" : ""} aria-label="Chọn ${esc(row.full_name)}"></td><td>${(model.page - 1) * state.studentPageSize + index + 1}</td><td><code>${esc(row.student_code)}</code></td><td class="wrap"><strong>${esc(row.full_name)}</strong><br><small>${fmtDate(row.birth_date)}</small></td><td>${esc(row.class_code || "—")}<br><small>${esc(campusName(row.campus_id))}</small></td><td><span class="badge blue">${esc(organizationLabel(row.organization_status))}</span></td><td class="wrap">${row.policy_groups ? `<span class="badge yellow">Chính sách</span> ${esc(row.policy_groups)}` : ""}${!["", "none"].includes(row.difficulty_status || "none") ? `<br><span class="badge red">${esc(difficultyLabel(row.difficulty_status))}</span>` : ""}${row.special_needs ? `<br><small>${esc(row.special_needs)}</small>` : ""}${!row.policy_groups && ["", "none"].includes(row.difficulty_status || "none") && !row.special_needs ? "—" : ""}</td><td>${row.status === "inactive" ? '<span class="badge">Ngừng học</span>' : row.status === "transferred" ? '<span class="badge yellow">Chuyển trường</span>' : row.status === "graduated" ? '<span class="badge blue">Hoàn thành</span>' : '<span class="badge green">Đang học</span>'}</td><td><button class="link-btn" data-student-view="${row.id}">Xem</button><button class="link-btn" data-student-edit="${row.id}">Sửa</button><button class="link-btn danger" data-student-delete="${row.id}">Xóa/Ngừng</button></td></tr>`,
            )
            .join("") || '<tr><td colspan="9" class="empty">Chưa có học sinh theo bộ lọc.</td></tr>'}</tbody></table></div>${paginationHtml("student", model)}</div></div>`,
      );

      let searchTimer;
      $("#studentSearch").oninput = (event) => {
        clearTimeout(searchTimer);
        searchTimer = setTimeout(() => {
          state.studentQuery = event.target.value;
          state.studentPage = 1;
          renderStudents();
        }, 250);
      };
      $("#studentClassFilter").onchange = (event) => {
        state.studentClassFilter = event.target.value;
        state.studentPage = 1;
        renderStudents();
      };
      $("#studentStatusFilter").onchange = (event) => {
        state.studentStatusFilter = event.target.value;
        state.studentPage = 1;
        renderStudents();
      };
      $("#studentOrganizationFilter").onchange = (event) => {
        state.studentOrganizationFilter = event.target.value;
        state.studentPage = 1;
        renderStudents();
      };
      $("#studentSupportFilter").onchange = (event) => {
        state.studentSupportFilter = event.target.value;
        state.studentPage = 1;
        renderStudents();
      };
      $("#studentPageSize").onchange = (event) => {
        state.studentPageSize = Number(event.target.value);
        state.studentPage = 1;
        renderStudents();
      };
      $("#studentPrev").onclick = () => {
        state.studentPage -= 1;
        renderStudents();
      };
      $("#studentNext").onclick = () => {
        state.studentPage += 1;
        renderStudents();
      };
      $("#studentSelectPage").onchange = (event) => {
        for (const row of model.rows)
          if (event.target.checked) selected.add(row.id);
          else selected.delete(row.id);
        renderStudents();
      };
      $$('[data-student-select]').forEach((input) => {
        input.onchange = () => {
          if (input.checked) selected.add(input.dataset.studentSelect);
          else selected.delete(input.dataset.studentSelect);
          renderStudents();
        };
      });
      $$('[data-student-edit]').forEach(
        (button) => (button.onclick = () => openStudentForm(button.dataset.studentEdit)),
      );
      $$('[data-student-view]').forEach(
        (button) => (button.onclick = () => viewStudent(button.dataset.studentView)),
      );
      $$('[data-student-delete]').forEach(
        (button) =>
          (button.onclick = () =>
            deleteStudentSafely(button.dataset.studentDelete)),
      );
      $("#studentAdd").onclick = () => openStudentForm();
      $("#studentTemplate").onclick = () => downloadImportTemplate("students", true);
      $("#studentImport").onclick = () => {
        state.importType = "students";
        state.importSession = null;
        go("imports");
      };
      const columns = [
        { name: "student_code", label: "Mã học sinh" },
        { name: "full_name", label: "Họ và tên" },
        { name: "birth_date", label: "Ngày sinh" },
        { name: "gender", label: "Giới tính" },
        { name: "class_code", label: "Mã lớp" },
        { name: "campus_code", label: "Mã cơ sở" },
        { name: "school_year", label: "Năm học" },
        { name: "ordinal", label: "Số thứ tự" },
        { name: "status", label: "Trạng thái" },
        { name: "organization_status", label: "Tình trạng Đội" },
        { name: "organization_joined_date", label: "Ngày kết nạp" },
        { name: "policy_groups", label: "Diện chính sách" },
        { name: "difficulty_status", label: "Tình trạng khó khăn" },
        { name: "special_needs", label: "Đặc điểm cần lưu ý" },
        { name: "support_notes", label: "Hỗ trợ và theo dõi" },
        { name: "guardian_name", label: "Người liên hệ" },
        { name: "guardian_phone", label: "Điện thoại liên hệ" },
        { name: "notes", label: "Ghi chú" },
      ];
      $("#studentCsv").onclick = () => exportCsv(rows, columns, `hoc-sinh-${today()}`);
      $("#studentXlsx").onclick = () =>
        exportXlsx(rows, columns, `hoc-sinh-${today()}`, "DANH SÁCH HỌC SINH").catch((error) => toast(error.message, "bad"));
      $("#studentWord").onclick = () =>
        exportWordTable(rows, columns, `hoc-sinh-${today()}`, "DANH SÁCH HỌC SINH");
      if ($("#studentClearSelection"))
        $("#studentClearSelection").onclick = () => {
          selected.clear();
          renderStudents();
        };
      if ($("#studentBulkTransfer"))
        $("#studentBulkTransfer").onclick = () => bulkTransferStudents();
      if ($("#studentBulkInactive"))
        $("#studentBulkInactive").onclick = () => bulkStudentStatus("inactive");
      if ($("#studentBulkRestore"))
        $("#studentBulkRestore").onclick = () => bulkStudentStatus("active");
    }

    async function openStudentForm(id = null) {
      const [student, classes] = await Promise.all([
        id ? db.get("students", id) : Promise.resolve(null),
        db.all("classes"),
      ]);
      const row = student || {
        status: "active",
        organization_status: "chua_xac_dinh",
        difficulty_status: "none",
        school_year_id: state.yearId,
        campus_id: state.campusId === "all" ? state.cache.campuses?.[0]?.id : state.campusId,
      };
      const available = classes
        .filter((item) => item.school_year_id === state.yearId && item.active !== false)
        .sort((a, b) => String(a.class_name).localeCompare(String(b.class_name), "vi", { numeric: true }));
      openModal(
        id ? "Sửa học sinh" : "Thêm học sinh",
        `<form id="studentForm"><div class="form-grid"><div class="field"><label class="required">Mã học sinh</label><input name="student_code" value="${esc(row.student_code || "")}" required maxlength="50"></div><div class="field"><label class="required">Họ và tên</label><input name="full_name" value="${esc(row.full_name || "")}" required maxlength="150"></div><div class="field"><label>Ngày sinh</label><input name="birth_date" type="date" value="${esc(row.birth_date || "")}"></div><div class="field"><label>Giới tính</label><select name="gender"><option value="">— Không ghi nhận —</option><option value="Nam" ${row.gender === "Nam" ? "selected" : ""}>Nam</option><option value="Nữ" ${row.gender === "Nữ" ? "selected" : ""}>Nữ</option><option value="Khác" ${row.gender === "Khác" ? "selected" : ""}>Khác</option></select></div><div class="field"><label class="required">Lớp</label><select name="class_id" required><option value="">— Chọn lớp —</option>${available.map((item) => `<option value="${item.id}" ${row.class_id === item.id ? "selected" : ""}>${esc(item.class_name)} • ${esc(campusName(item.campus_id))}</option>`).join("")}</select></div><div class="field"><label>Số thứ tự</label><input name="ordinal" type="number" min="1" max="100" value="${esc(row.ordinal || "")}"></div><div class="field"><label>Trạng thái học</label><select name="status"><option value="active" ${row.status === "active" ? "selected" : ""}>Đang học</option><option value="inactive" ${row.status === "inactive" ? "selected" : ""}>Ngừng học</option><option value="transferred" ${row.status === "transferred" ? "selected" : ""}>Chuyển trường</option><option value="graduated" ${row.status === "graduated" ? "selected" : ""}>Hoàn thành cấp học</option></select></div><div class="field"><label>Tình trạng Đội</label><select name="organization_status">${[["chua_xac_dinh","Chưa xác định"],["nhi_dong","Nhi đồng"],["du_bi","Dự bị đội viên"],["doi_vien","Đội viên"],["doan_vien","Đoàn viên"]].map(([value,label]) => `<option value="${value}" ${(row.organization_status || "chua_xac_dinh") === value ? "selected" : ""}>${label}</option>`).join("")}</select></div><div class="field"><label>Ngày kết nạp/công nhận</label><input name="organization_joined_date" type="date" value="${esc(row.organization_joined_date || "")}"></div><div class="field full"><label>Diện chính sách</label><input name="policy_groups" value="${esc(row.policy_groups || "")}" maxlength="500" placeholder="Ví dụ: hộ nghèo; con thương binh; mồ côi…"></div><div class="field"><label>Tình trạng khó khăn</label><select name="difficulty_status"><option value="none" ${(row.difficulty_status || "none") === "none" ? "selected" : ""}>Không</option><option value="monitor" ${row.difficulty_status === "monitor" ? "selected" : ""}>Cần theo dõi</option><option value="support" ${row.difficulty_status === "support" ? "selected" : ""}>Cần hỗ trợ</option><option value="urgent" ${row.difficulty_status === "urgent" ? "selected" : ""}>Cần hỗ trợ khẩn</option></select></div><div class="field"><label>Người liên hệ</label><input name="guardian_name" value="${esc(row.guardian_name || "")}" maxlength="150"></div><div class="field"><label>Điện thoại liên hệ</label><input name="guardian_phone" value="${esc(row.guardian_phone || "")}" inputmode="tel" maxlength="30"></div><div class="field full"><label>Đặc điểm cần lưu ý/học sinh đặc biệt</label><textarea name="special_needs" maxlength="1500">${esc(row.special_needs || "")}</textarea></div><div class="field full"><label>Nội dung hỗ trợ và theo dõi</label><textarea name="support_notes" maxlength="2000">${esc(row.support_notes || "")}</textarea></div><div class="field full"><label>Ghi chú chung</label><textarea name="notes" maxlength="1000">${esc(row.notes || "")}</textarea></div></div></form>`,
        `<button class="btn" id="cancelStudent">Hủy</button><button class="btn primary" id="saveStudent">Lưu học sinh</button>`,
        true,
      );
      $("#cancelStudent").onclick = closeModal;
      $("#saveStudent").onclick = async () => {
        const form = $("#studentForm");
        if (!form.reportValidity()) return;
        const data = Object.fromEntries(new FormData(form));
        data.student_code = engine.codeText(data.student_code);
        const duplicate = (await db.all("students")).find(
          (item) =>
            item.id !== id &&
            item.school_year_id === state.yearId &&
            engine.codeText(item.student_code) === data.student_code,
        );
        if (duplicate) return toast("Mã học sinh đã tồn tại trong năm học.", "bad");
        const klass = available.find((item) => item.id === data.class_id);
        if (!klass) return toast("Lớp không hợp lệ.", "bad");
        const saved = await db.put("students", {
          ...row,
          ...data,
          ordinal: data.ordinal ? Number(data.ordinal) : null,
          school_year_id: state.yearId,
          academic_year_id: state.yearId,
          class_id: klass.id,
          class_code: klass.code || klass.class_name,
          class_name: klass.class_name,
          campus_id: klass.campus_id,
          campus_code: state.cache.campuses?.find((item) => item.id === klass.campus_id)?.code || "",
        });
        await propagateStudentIdentity(saved, student);
        closeModal();
        toast("Đã lưu học sinh");
        renderStudents();
      };
    }

    async function propagateStudentIdentity(student, previous = null) {
      if (!student?.id) return;
      const codes = new Set(
        [student.student_code, previous?.student_code]
          .filter(Boolean)
          .map(engine.codeText),
      );
      const matches = (row) =>
        row.student_id === student.id ||
        (row.student_code && codes.has(engine.codeText(row.student_code)));
      const updates = {
        student_incidents: (row) => ({
          ...row,
          student_id: student.id,
          student_code: student.student_code,
          student_name: student.full_name,
          class_id: student.class_id,
          class_code: student.class_code,
          campus_id: student.campus_id,
          campus_code: student.campus_code,
        }),
        program_results: (row) => ({
          ...row,
          student_id: student.id,
          student_code: student.student_code,
          student_name: student.full_name,
          class_id: student.class_id,
          class_code: student.class_code,
          campus_id: student.campus_id,
        }),
        team_members: (row) => ({
          ...row,
          student_id: student.id,
          student_code: student.student_code,
          name: student.full_name,
          class_id: student.class_id,
          class_name: student.class_name,
          campus_id: student.campus_id,
        }),
        commendations: (row) => ({
          ...row,
          student_id: student.id,
          student_code: student.student_code,
          recipient: student.full_name,
          class_id: student.class_id,
          class_name: student.class_name,
          campus_id: student.campus_id,
        }),
      };
      for (const [store, transform] of Object.entries(updates)) {
        for (const row of (await db.all(store)).filter(matches))
          await db.put(store, transform(row));
      }
    }

    async function deleteStudentSafely(id) {
      const student = await db.get("students", id);
      if (!student) return;
      const dependentStores = [
          "student_incidents",
          "program_results",
          "team_members",
          "commendations",
        ],
        dependencies = [];
      for (const store of dependentStores) {
        const count = (await db.all(store)).filter(
          (row) =>
            row.student_id === id ||
            (row.student_code &&
              engine.codeText(row.student_code) ===
                engine.codeText(student.student_code)),
        ).length;
        if (count) dependencies.push(`${store}: ${count}`);
      }
      if (dependencies.length) {
        if (
          !confirm(
            `Học sinh “${student.full_name}” đã có dữ liệu liên quan (${dependencies.join(", ")}). Không xóa lịch sử; chuyển sang trạng thái Ngừng học?`,
          )
        )
          return;
        await db.put("students", {
          ...student,
          status: "inactive",
          inactive_reason: "Người dùng chọn Xóa/Ngừng khi đã có dữ liệu liên quan",
          inactive_at: now(),
        });
        toast("Đã ngừng học; toàn bộ lịch sử liên quan được giữ nguyên.");
      } else {
        if (!confirm(`Xóa học sinh “${student.full_name}” chưa có dữ liệu liên quan?`))
          return;
        await db.remove("students", id);
        toast("Đã xóa học sinh chưa phát sinh dữ liệu; có dấu vết đồng bộ.");
      }
      state.studentSelected.delete(id);
      renderStudents();
    }

    async function viewStudent(id) {
      const [student, incidents, commendations, programs, links] = await Promise.all([
        db.get("students", id),
        db.all("student_incidents"),
        db.all("commendations"),
        db.all("program_results"),
        db.all("document_links"),
      ]);
      if (!student) return;
      const relatedIncidents = incidents.filter(
        (row) => row.student_id === id || engine.codeText(row.student_code) === engine.codeText(student.student_code),
      );
      const relatedAwards = commendations.filter((row) =>
        normalizeText(row.recipient).includes(normalizeText(student.full_name)),
      );
      const relatedPrograms = programs.filter(
        (row) =>
          row.student_id === id ||
          engine.codeText(row.student_code) === engine.codeText(student.student_code) ||
          normalizeText(row.scope || "").includes(normalizeText(student.full_name)),
      );
      const evidenceCount = links.filter(
        (row) =>
          ["students", "student_incidents"].includes(row.related_module) &&
          [id, student.student_code].includes(row.related_record_id),
      ).length;
      openModal(
        "Chi tiết học sinh",
        `<div class="grid-2"><div class="card"><div class="card-body"><div class="split"><span>Mã học sinh</span><code>${esc(student.student_code)}</code></div><div class="split mt"><span>Họ và tên</span><strong>${esc(student.full_name)}</strong></div><div class="split mt"><span>Ngày sinh</span><span>${fmtDate(student.birth_date)}</span></div><div class="split mt"><span>Lớp</span><span>${esc(student.class_code || student.class_name || "—")}</span></div><div class="split mt"><span>Cơ sở</span><span>${esc(campusName(student.campus_id))}</span></div><div class="split mt"><span>Tình trạng Đội</span><strong>${esc(organizationLabel(student.organization_status))}</strong></div><div class="split mt"><span>Ngày kết nạp/công nhận</span><span>${fmtDate(student.organization_joined_date)}</span></div></div></div><div class="card"><div class="card-body"><div class="metric-row" style="grid-template-columns:1fr 1fr"><div class="metric"><strong>${relatedIncidents.length}</strong><span>Vi phạm/ghi nhận</span></div><div class="metric"><strong>${relatedAwards.length}</strong><span>Khen thưởng</span></div></div><div class="split mt"><span>Diện chính sách</span><strong>${esc(student.policy_groups || "Không ghi nhận")}</strong></div><div class="split mt"><span>Khó khăn</span><strong>${esc(difficultyLabel(student.difficulty_status))}</strong></div><div class="split mt"><span>Người liên hệ</span><span>${esc([student.guardian_name, student.guardian_phone].filter(Boolean).join(" • ") || "—")}</span></div></div></div></div><div class="card mt"><div class="card-head"><h2>Hỗ trợ và lưu ý</h2></div><div class="card-body"><p><strong>Đặc điểm cần lưu ý:</strong> ${esc(student.special_needs || "Chưa ghi nhận")}</p><p><strong>Nội dung hỗ trợ:</strong> ${esc(student.support_notes || "Chưa ghi nhận")}</p><p><strong>Ghi chú:</strong> ${esc(student.notes || "—")}</p></div></div><div class="card mt"><div class="card-head"><h2>Lịch sử vi phạm gần nhất</h2></div><div class="card-body">${relatedIncidents.length ? relatedIncidents.slice(-10).reverse().map((row) => `<div class="split"><span>${fmtDate(row.date)} • ${esc(row.incident_type)}</span><span>${esc(row.status || "draft")}</span></div>`).join("") : '<div class="empty">Chưa có vi phạm.</div>'}</div></div>`,
        `<button class="btn" id="studentViewClose">Đóng</button><button class="btn primary" id="studentViewIncident">＋ Ghi nhận vi phạm</button>`,
        true,
      );
      $("#modalBody").insertAdjacentHTML(
        "beforeend",
        `<div class="card mt"><div class="card-head"><h2>Rèn luyện và minh chứng</h2></div><div class="card-body"><div class="split"><span>Kết quả rèn luyện/phong trào liên quan</span><strong>${relatedPrograms.length}</strong></div><div class="split mt"><span>Hồ sơ minh chứng được gắn trực tiếp</span><strong>${evidenceCount}</strong></div>${relatedPrograms.length ? `<div class="mt">${relatedPrograms.slice(-8).reverse().map((row) => `<div class="split"><span>${esc(row.name || "Chương trình")}</span><span>${esc(row.status || row.result || "Đang theo dõi")}</span></div>`).join("")}</div>` : ""}</div></div>`,
      );
      $("#modalFoot").insertAdjacentHTML(
        "beforeend",
        '<input id="studentEvidenceFiles" type="file" accept="image/png,image/jpeg,image/webp,application/pdf" multiple class="hidden"><button class="btn" id="studentViewEvidence">Gắn minh chứng</button>',
      );
      $("#studentViewEvidence").onclick = () => $("#studentEvidenceFiles").click();
      $("#studentEvidenceFiles").onchange = async (event) => {
        const files = [...event.target.files];
        event.target.value = "";
        if (!files.length) return;
        await saveDocumentFiles(
          files,
          files.map((file) => ({
            name: file.name.replace(/\.[^.]+$/, ""),
            category: "Minh chứng",
            related_module: "students",
            related_record_id: student.id,
          })),
        );
        closeModal();
        toast(`Đã gắn ${files.length} tệp minh chứng với học sinh.`);
      };
      $("#studentViewClose").onclick = closeModal;
      $("#studentViewIncident").onclick = () => {
        closeModal();
        openIncidentForm(null, student);
      };
    }

    async function bulkTransferStudents() {
      const classes = (await db.all("classes"))
        .filter((row) => row.school_year_id === state.yearId && row.active !== false)
        .sort((a, b) => String(a.class_name).localeCompare(String(b.class_name), "vi", { numeric: true }));
      openModal(
        `Chuyển ${state.studentSelected.size} học sinh`,
        `<div class="notice warn">Hệ thống sẽ giữ lịch sử vi phạm, khen thưởng và mã học sinh; chỉ cập nhật lớp hiện hành.</div><div class="field mt"><label class="required">Lớp đích</label><select id="studentTargetClass"><option value="">— Chọn lớp —</option>${classes.map((row) => `<option value="${row.id}">${esc(row.class_name)} • ${esc(campusName(row.campus_id))}</option>`).join("")}</select></div>`,
        `<button class="btn" id="cancelStudentTransfer">Hủy</button><button class="btn primary" id="confirmStudentTransfer">Xác nhận chuyển lớp</button>`,
      );
      $("#cancelStudentTransfer").onclick = closeModal;
      $("#confirmStudentTransfer").onclick = async () => {
        const klass = classes.find((row) => row.id === $("#studentTargetClass").value);
        if (!klass) return toast("Hãy chọn lớp đích.", "bad");
        const selectedRows = (await db.all("students")).filter((row) =>
          state.studentSelected.has(row.id),
        );
        await createInternalSnapshot("Trước chuyển lớp hàng loạt", {
          tier: "protected",
          protectedSnapshot: true,
          reason: "before-bulk-student-transfer",
          yearId: state.yearId,
        });
        const campus = state.cache.campuses?.find((row) => row.id === klass.campus_id);
        await db.bulkPut(
          "students",
          selectedRows.map((row) => ({
            ...row,
            class_id: klass.id,
            class_code: klass.code || klass.class_name,
            class_name: klass.class_name,
            campus_id: klass.campus_id,
            campus_code: campus?.code || "",
          })),
        );
        state.studentSelected.clear();
        closeModal();
        toast(`Đã chuyển ${selectedRows.length} học sinh sang ${klass.class_name}.`);
        renderStudents();
      };
    }

    async function bulkStudentStatus(status) {
      const selectedRows = (await db.all("students")).filter((row) =>
        state.studentSelected.has(row.id),
      );
      if (!selectedRows.length) return;
      openModal(
        status === "active" ? "Khôi phục học sinh" : "Ngừng học hàng loạt",
        `<div class="notice warn">Sẽ cập nhật ${selectedRows.length} học sinh. Dữ liệu lịch sử được giữ nguyên.</div>`,
        `<button class="btn" id="cancelStudentStatus">Hủy</button><button class="btn primary" id="confirmStudentStatus">Xác nhận</button>`,
      );
      $("#cancelStudentStatus").onclick = closeModal;
      $("#confirmStudentStatus").onclick = async () => {
        await createInternalSnapshot("Trước đổi trạng thái học sinh hàng loạt", {
          tier: "protected",
          protectedSnapshot: true,
          reason: "before-bulk-student-status",
          yearId: state.yearId,
        });
        await db.bulkPut(
          "students",
          selectedRows.map((row) => ({ ...row, status })),
        );
        state.studentSelected.clear();
        closeModal();
        toast(`Đã cập nhật ${selectedRows.length} học sinh.`);
        renderStudents();
      };
    }

    async function incidentsForView() {
      const rows = await db.all("student_incidents");
      const query = normalizeText(state.incidentQuery);
      return rows
        .filter((row) => !state.yearId || row.school_year_id === state.yearId)
        .filter((row) => state.campusId === "all" || row.campus_id === state.campusId)
        .filter(
          (row) =>
            state.incidentClassFilter === "all" ||
            row.class_id === state.incidentClassFilter,
        )
        .filter(
          (row) =>
            state.incidentStatusFilter === "all" ||
            (row.status || "draft") === state.incidentStatusFilter,
        )
        .filter((row) => {
          if (!query) return true;
          return normalizeText(
            [
              row.incident_code,
              row.student_code,
              row.student_name,
              row.class_code,
              row.incident_type,
              row.content,
              row.recorded_by,
            ]
              .filter(Boolean)
              .join(" "),
          ).includes(query);
        })
        .sort((a, b) => String(b.date || b.created_at).localeCompare(String(a.date || a.created_at)));
    }

    async function renderIncidents() {
      const [rows, classes] = await Promise.all([
        incidentsForView(),
        db.all("classes"),
      ]);
      const availableClasses = classes
        .filter((row) => row.school_year_id === state.yearId)
        .filter((row) => state.campusId === "all" || row.campus_id === state.campusId)
        .sort((a, b) => String(a.class_name).localeCompare(String(b.class_name), "vi", { numeric: true }));
      const model = pageSlice(rows, state.incidentPage, state.incidentPageSize);
      state.incidentPage = model.page;
      const confirmed = rows.filter((row) => row.status === "confirmed").length;
      const proposed = rows.reduce((sum, row) => sum + (Number(row.proposed_deduction) || 0), 0);
      setContent(
        pageHead(
          "Vi phạm học sinh",
          "Sổ ghi nhận có cấu trúc; điểm trừ chỉ là đề xuất và không tự thay đổi bảng thi đua.",
          `<button class="btn" id="incidentTemplate">Tải mẫu Excel</button><button class="btn" id="incidentImport">Nhập dữ liệu</button><button class="btn primary" id="incidentAdd">＋ Ghi nhận</button>`,
        ) +
          `<div class="metric-row"><div class="metric"><strong>${rows.length}</strong><span>Tổng ghi nhận</span></div><div class="metric"><strong>${confirmed}</strong><span>Đã xác nhận</span></div><div class="metric"><strong>${rows.filter((row) => row.status === "draft").length}</strong><span>Bản nháp</span></div><div class="metric"><strong>${new Set(rows.map((row) => row.student_code).filter(Boolean)).size}</strong><span>Học sinh liên quan</span></div><div class="metric"><strong>${proposed.toLocaleString("vi-VN")}</strong><span>Điểm trừ đề xuất</span></div></div>
          <div class="card mt"><div class="card-body"><div class="toolbar"><input class="grow" id="incidentSearch" value="${esc(state.incidentQuery)}" placeholder="Tìm học sinh, lớp, loại hoặc nội dung…"><select id="incidentClassFilter"><option value="all">Tất cả lớp</option>${availableClasses.map((row) => `<option value="${row.id}" ${state.incidentClassFilter === row.id ? "selected" : ""}>${esc(row.class_name)}</option>`).join("")}</select><select id="incidentStatusFilter"><option value="all">Tất cả trạng thái</option><option value="draft" ${state.incidentStatusFilter === "draft" ? "selected" : ""}>Bản nháp</option><option value="confirmed" ${state.incidentStatusFilter === "confirmed" ? "selected" : ""}>Đã xác nhận</option><option value="resolved" ${state.incidentStatusFilter === "resolved" ? "selected" : ""}>Đã xử lý</option></select><button class="btn small" id="incidentCsv">CSV</button><button class="btn small" id="incidentXlsx">Excel</button><button class="btn small" id="incidentWord">Word</button></div><div class="notice warn">Không tự trừ điểm và không thay đổi bảng thi đua đã khóa. Hãy duyệt quy tắc thi đua riêng.</div><div class="table-wrap" style="max-height:520px"><table><thead><tr><th>Ngày</th><th>Mã</th><th>Học sinh</th><th>Lớp</th><th>Loại</th><th>Nội dung</th><th>Đề xuất</th><th>Trạng thái</th><th></th></tr></thead><tbody>${model.rows
            .map(
              (row) =>
                `<tr><td>${fmtDate(row.date)}</td><td><code>${esc(row.incident_code)}</code></td><td class="wrap"><strong>${esc(row.student_name || "—")}</strong><br><small>${esc(row.student_code || "")}</small></td><td>${esc(row.class_code || "—")}</td><td>${esc(row.incident_type || "—")}</td><td class="wrap">${esc(row.content || "")}</td><td>${Number(row.proposed_deduction || 0).toLocaleString("vi-VN")}</td><td>${row.status === "confirmed" ? '<span class="badge green">Đã xác nhận</span>' : row.status === "resolved" ? '<span class="badge blue">Đã xử lý</span>' : '<span class="badge yellow">Bản nháp</span>'}</td><td><button class="link-btn" data-incident-edit="${row.id}">Sửa</button></td></tr>`,
            )
            .join("") || '<tr><td colspan="9" class="empty">Chưa có ghi nhận theo bộ lọc.</td></tr>'}</tbody></table></div>${paginationHtml("incident", model)}</div></div>`,
      );
      let timer;
      $("#incidentSearch").oninput = (event) => {
        clearTimeout(timer);
        timer = setTimeout(() => {
          state.incidentQuery = event.target.value;
          state.incidentPage = 1;
          renderIncidents();
        }, 250);
      };
      $("#incidentClassFilter").onchange = (event) => {
        state.incidentClassFilter = event.target.value;
        state.incidentPage = 1;
        renderIncidents();
      };
      $("#incidentStatusFilter").onchange = (event) => {
        state.incidentStatusFilter = event.target.value;
        state.incidentPage = 1;
        renderIncidents();
      };
      $("#incidentPrev").onclick = () => {
        state.incidentPage -= 1;
        renderIncidents();
      };
      $("#incidentNext").onclick = () => {
        state.incidentPage += 1;
        renderIncidents();
      };
      $("#incidentAdd").onclick = () => openIncidentForm();
      $$('[data-incident-edit]').forEach(
        (button) => (button.onclick = () => openIncidentForm(button.dataset.incidentEdit)),
      );
      $("#incidentTemplate").onclick = () => downloadImportTemplate("student_incidents", true);
      $("#incidentImport").onclick = () => {
        state.importType = "student_incidents";
        state.importSession = null;
        go("imports");
      };
      const columns = [
        { name: "incident_code", label: "Mã vi phạm" },
        { name: "date", label: "Ngày" },
        { name: "week", label: "Tuần" },
        { name: "student_code", label: "Mã học sinh" },
        { name: "student_name", label: "Họ tên" },
        { name: "class_code", label: "Mã lớp" },
        { name: "campus_code", label: "Mã cơ sở" },
        { name: "incident_type", label: "Loại vi phạm" },
        { name: "content", label: "Nội dung" },
        { name: "proposed_deduction", label: "Điểm trừ đề xuất" },
        { name: "recorded_by", label: "Người ghi nhận" },
        { name: "status", label: "Trạng thái" },
      ];
      $("#incidentCsv").onclick = () => exportCsv(rows, columns, `vi-pham-${today()}`);
      $("#incidentXlsx").onclick = () =>
        exportXlsx(rows, columns, `vi-pham-${today()}`, "TỔNG HỢP VI PHẠM HỌC SINH").catch((error) => toast(error.message, "bad"));
      $("#incidentWord").onclick = () =>
        exportWordTable(rows, columns, `vi-pham-${today()}`, "TỔNG HỢP VI PHẠM HỌC SINH");
    }

    async function openIncidentForm(id = null, studentPreset = null) {
      const [incident, classes, students] = await Promise.all([
        id ? db.get("student_incidents", id) : Promise.resolve(null),
        db.all("classes"),
        db.all("students"),
      ]);
      const row = incident || {
        incident_code: `VP-${today().replace(/-/g, "")}-${String(Date.now()).slice(-5)}`,
        date: today(),
        week_id: state.weekId || "",
        status: "draft",
        student_id: studentPreset?.id || "",
        student_code: studentPreset?.student_code || "",
        student_name: studentPreset?.full_name || "",
        class_id: studentPreset?.class_id || "",
        class_code: studentPreset?.class_code || "",
        campus_id: studentPreset?.campus_id || (state.campusId === "all" ? "" : state.campusId),
      };
      const availableClasses = classes
        .filter((item) => item.school_year_id === state.yearId)
        .sort((a, b) => String(a.class_name).localeCompare(String(b.class_name), "vi", { numeric: true }));
      const studentList = students.filter((item) => item.school_year_id === state.yearId);
      openModal(
        id ? "Sửa vi phạm học sinh" : "Ghi nhận vi phạm học sinh",
        `<form id="structuredIncidentForm"><div class="form-grid"><div class="field"><label class="required">Mã vi phạm</label><input name="incident_code" value="${esc(row.incident_code)}" required></div><div class="field"><label class="required">Ngày</label><input name="date" type="date" value="${esc(row.date || today())}" required></div><div class="field"><label>Tuần</label><input name="week" type="number" min="1" max="53" value="${esc(row.week || "")}"></div><div class="field"><label class="required">Mã học sinh</label><input name="student_code" list="incidentStudentCodes" value="${esc(row.student_code || "")}" required><datalist id="incidentStudentCodes">${studentList.slice(0, 5000).map((item) => `<option value="${esc(item.student_code)}">${esc(item.full_name)} • ${esc(item.class_code)}</option>`).join("")}</datalist></div><div class="field"><label class="required">Họ tên</label><input name="student_name" value="${esc(row.student_name || "")}" required></div><div class="field"><label class="required">Lớp</label><select name="class_id" required><option value="">— Chọn lớp —</option>${availableClasses.map((item) => `<option value="${item.id}" ${row.class_id === item.id ? "selected" : ""}>${esc(item.class_name)} • ${esc(campusName(item.campus_id))}</option>`).join("")}</select></div><div class="field"><label class="required">Loại vi phạm</label><input name="incident_type" value="${esc(row.incident_type || "")}" required maxlength="100"></div><div class="field"><label>Điểm trừ đề xuất</label><input name="proposed_deduction" type="number" step="0.5" min="0" value="${esc(row.proposed_deduction ?? "")}"></div><div class="field full"><label class="required">Nội dung</label><textarea name="content" required maxlength="2000">${esc(row.content || "")}</textarea></div><div class="field"><label>Minh chứng/mã hồ sơ</label><input name="evidence" value="${esc(row.evidence || "")}"></div><div class="field"><label>Người ghi nhận</label><input name="recorded_by" value="${esc(row.recorded_by || "")}"></div><div class="field"><label>Trạng thái</label><select name="status"><option value="draft" ${row.status === "draft" ? "selected" : ""}>Bản nháp</option><option value="confirmed" ${row.status === "confirmed" ? "selected" : ""}>Đã xác nhận</option><option value="resolved" ${row.status === "resolved" ? "selected" : ""}>Đã xử lý</option></select></div></div></form><div class="notice warn mt">Việc xác nhận không tự động trừ điểm và không mở khóa bảng thi đua.</div>`,
        `<button class="btn" id="cancelStructuredIncident">Hủy</button><button class="btn primary" id="saveStructuredIncident">Lưu ghi nhận</button>`,
        true,
      );
      $("#structuredIncidentForm .form-grid").insertAdjacentHTML(
        "beforeend",
        '<div class="field full"><label>Ảnh/PDF minh chứng</label><input id="incidentEvidenceFiles" type="file" accept="image/png,image/jpeg,image/webp,application/pdf" multiple><small>Tệp gốc được lưu trong Hồ sơ và gắn với vi phạm sau khi lưu.</small></div>',
      );
      $("#cancelStructuredIncident").onclick = closeModal;
      const codeInput = $('#structuredIncidentForm [name="student_code"]');
      codeInput.onchange = () => {
        const student = studentList.find(
          (item) => engine.codeText(item.student_code) === engine.codeText(codeInput.value),
        );
        if (!student) return;
        $('#structuredIncidentForm [name="student_name"]').value = student.full_name || "";
        $('#structuredIncidentForm [name="class_id"]').value = student.class_id || "";
      };
      $("#saveStructuredIncident").onclick = async () => {
        const form = $("#structuredIncidentForm");
        if (!form.reportValidity()) return;
        const data = Object.fromEntries(new FormData(form));
        data.incident_code = engine.codeText(data.incident_code);
        data.student_code = engine.codeText(data.student_code);
        const duplicate = (await db.all("student_incidents")).find(
          (item) => item.id !== id && engine.codeText(item.incident_code) === data.incident_code,
        );
        if (duplicate) return toast("Mã vi phạm đã tồn tại.", "bad");
        const klass = availableClasses.find((item) => item.id === data.class_id);
        if (!klass) return toast("Lớp không hợp lệ.", "bad");
        const student = studentList.find(
          (item) => engine.codeText(item.student_code) === data.student_code,
        );
        const savedIncident = await db.put("student_incidents", {
          ...row,
          ...data,
          proposed_deduction: data.proposed_deduction === "" ? null : Number(data.proposed_deduction),
          week: data.week === "" ? null : Number(data.week),
          school_year_id: state.yearId,
          academic_year_id: state.yearId,
          class_id: klass.id,
          class_code: klass.code || klass.class_name,
          campus_id: klass.campus_id,
          campus_code: state.cache.campuses?.find((item) => item.id === klass.campus_id)?.code || "",
          student_id: student?.id || row.student_id || null,
        });
        const evidenceFiles = [...($("#incidentEvidenceFiles")?.files || [])];
        if (evidenceFiles.length)
          await saveDocumentFiles(
            evidenceFiles,
            evidenceFiles.map((file) => ({
              name: file.name.replace(/\.[^.]+$/, ""),
              category: "Minh chứng",
              related_module: "student_incidents",
              related_record_id: savedIncident.id,
            })),
          );
        closeModal();
        toast("Đã lưu ghi nhận; bảng thi đua không bị thay đổi.");
        if (state.page === "incidents") renderIncidents();
      };
    }

    const TEMPLATE_COLUMNS = {
      campuses: ["MaCoSo", "TenCoSo", "DiaChi", "TrangThai", "GhiChu"],
      classes: [
        "MaLop",
        "TenLop",
        "Khoi",
        "MaCoSo",
        "TenCoSo",
        "GiaoVienChuNhiem",
        "SiSo",
        "TrangThai",
        "GhiChu",
      ],
      students: [
        "MaHocSinh",
        "HoVaTen",
        "NgaySinh",
        "GioiTinh",
        "MaLop",
        "TenLop",
        "MaCoSo",
        "NamHoc",
        "SoThuTu",
        "TrangThai",
        "TinhTrangDoi",
        "NgayKetNap",
        "DienChinhSach",
        "TinhTrangKhoKhan",
        "HocSinhDacBiet",
        "HoTroTheoDoi",
        "NguoiLienHe",
        "DienThoaiLienHe",
        "GhiChu",
      ],
      plans: [
        "MaKeHoach",
        "TenKeHoach",
        "CapKeHoach",
        "NgayBatDau",
        "NgayKetThuc",
        "MucTieu",
        "TrangThai",
        "TienDo",
        "KetQua",
        "MinhChung",
        "GhiChu",
      ],
      tasks: [
        "MaCongViec",
        "TenCongViec",
        "MoTa",
        "MaKeHoach",
        "NgayBatDau",
        "HanHoanThanh",
        "UuTien",
        "NguoiPhuTrach",
        "MaCoSo",
        "TrangThai",
        "TienDo",
        "NgayHoanThanh",
        "KetQua",
        "MinhChung",
        "GhiChu",
      ],
      task_check_items: [
        "MaCongViec",
        "NoiDung",
        "BatBuoc",
        "DaHoanThanh",
        "ThuTu",
      ],
      calendar_events: [
        "MaSuKien",
        "TenSuKien",
        "Ngay",
        "GioBatDau",
        "GioKetThuc",
        "Loai",
        "DiaDiem",
        "TrangThai",
      ],
      activities: [
        "MaHoatDong",
        "TenHoatDong",
        "NhomHoatDong",
        "NgayToChuc",
        "DiaDiem",
        "NguoiPhuTrach",
        "DoiTuong",
        "TrangThai",
      ],
      criteria: ["MaTieuChi", "TenTieuChi", "Nhom", "DiemToiDa", "DangSuDung"],
      score_entries: ["MaLop", "MaTieuChi", "Tuan", "Diem", "GhiChu"],
      student_incidents: [
        "MaViPham",
        "Ngay",
        "Tuan",
        "MaHocSinh",
        "HoTen",
        "MaLop",
        "MaCoSo",
        "LoaiViPham",
        "NoiDung",
        "DiemTruDeXuat",
        "MinhChung",
        "NguoiGhiNhan",
        "TrangThai",
      ],
      team_members: ["Ma", "HoVaTen", "MaLop", "DoiBan", "ChucVu", "NhiemKy"],
      program_results: [
        "Ma",
        "TenChuongTrinh",
        "DoiTuong",
        "KetQua",
        "NgayCongNhan",
        "TrangThai",
      ],
      commendations: [
        "Ma",
        "LoaiKhenThuong",
        "DoiTuong",
        "ThanhTich",
        "Ngay",
        "TrangThai",
      ],
      documents: [
        "MaHoSo",
        "TenHoSo",
        "Loai",
        "Ngay",
        "DonViBanHanh",
        "MoTa",
        "TrangThai",
      ],
      equipment: [
        "MaThietBi",
        "TenThietBi",
        "Nhom",
        "SoLuong",
        "DonViTinh",
        "TinhTrang",
        "NoiLuu",
        "GhiChu",
      ],
      equipment_transactions: [
        "MaGiaoNhan",
        "MaThietBi",
        "Ngay",
        "Loai",
        "SoLuong",
        "NguoiNhan",
        "GhiChu",
      ],
      generated_reports: [
        "MaBaoCao",
        "TenBaoCao",
        "LoaiBaoCao",
        "PhamVi",
        "TrangThai",
        "GhiChu",
      ],
    };

    const TEMPLATE_EXAMPLES = {
      campuses: ["CS1", "Cơ sở chính", "Phường 1", "active", "Dữ liệu minh họa"],
      classes: ["6A1", "6/1", "6", "CS1", "Cơ sở chính", "Nguyễn Văn A", "45", "active", "Dữ liệu minh họa"],
      students: ["HS0001", "Nguyễn Minh An", "2014-03-12", "Nam", "6A1", "6/1", "CS1", "2026-2027", "1", "active", "doi_vien", "2026-09-15", "Hộ cận nghèo", "monitor", "Cần lưu ý sức khỏe", "Theo dõi và phối hợp GVCN", "Nguyễn Văn B", "0900000000", "Dữ liệu minh họa"],
      plans: ["KH001", "Kế hoạch tháng 9", "month", "2026-09-01", "2026-09-30", "Tổ chức hoạt động đầu năm", "active", "50", "Đã triển khai 2/4 nội dung", "Hồ sơ/KH001", "Tiếp tục trong tuần sau"],
      tasks: ["CV001", "Chuẩn bị sinh hoạt dưới cờ", "Kiểm tra âm thanh", "KH001", "2026-09-01", "2026-09-05", "Cao", "Tổng phụ trách", "CS1", "done", "100", "2026-09-05", "Đã hoàn thành", "Ảnh sinh hoạt tuần 1", "Đã đối chiếu"],
      task_check_items: ["CV001", "Kiểm tra âm thanh", "true", "false", "1"],
      calendar_events: ["SK001", "Sinh hoạt dưới cờ", "2026-09-07", "07:00", "07:45", "Hoạt động Đội", "Sân trường", "planned"],
      activities: ["HD001", "Ngày hội thiếu nhi", "Văn nghệ – thể thao", "2026-10-15", "Sân trường", "Tổng phụ trách", "Toàn trường", "planned"],
      criteria: ["TC01", "Nề nếp", "Nề nếp", "10", "true"],
      score_entries: ["6A1", "TC01", "1", "9.5", "Dữ liệu minh họa"],
      student_incidents: ["VP0001", "2026-09-08", "1", "HS0001", "Nguyễn Minh An", "6A1", "CS1", "Đi học muộn", "Đến lớp sau giờ quy định", "1", "", "Tổng phụ trách", "draft"],
      team_members: ["TV001", "Nguyễn Minh An", "6A1", "Ban Chỉ huy Liên đội", "Ủy viên", "2026-2027"],
      program_results: ["RL001", "Rèn luyện đội viên", "6A1", "Đang theo dõi", "", "draft"],
      commendations: ["KT001", "Cá nhân", "Nguyễn Minh An", "Tham gia tích cực", "2026-11-20", "draft"],
      documents: ["HS001", "Kế hoạch tháng 9", "Kế hoạch", "2026-09-01", "Nhà trường", "Dữ liệu minh họa", "draft"],
      equipment: ["TB001", "Trống Đội", "Nghi lễ", "2", "Cái", "Tốt", "Kho Đội", "Dữ liệu minh họa"],
      equipment_transactions: ["GN001", "TB001", "2026-09-01", "Mượn", "1", "Nguyễn Văn A", "Dữ liệu minh họa"],
      generated_reports: ["BC001", "Báo cáo công tác tuần 1", "Công tác tuần", "Tuần 1", "draft", "Dữ liệu minh họa"],
    };

    return {
      migrateSmartImportData,
      renderStudents,
      renderIncidents,
      openIncidentForm,
      renderImportCenter: () => renderImportCenter(),
      renderClassManagement: (panel) => renderClassManagement(panel),
      downloadImportTemplate,
      exportRows: async (rows, columns, fileName, title, format) => {
        if (format === "xlsx") return exportXlsx(rows, columns, fileName, title);
        if (format === "word") return exportWordTable(rows, columns, fileName, title);
        return exportCsv(rows, columns, fileName);
      },
    };

    function newImportSession(type = state.importType || "classes") {
      return {
        id: uid(),
        type,
        status: "idle",
        sourceName: "",
        sourceChecksum: "",
        files: [],
        tables: [],
        tableIndex: 0,
        headerIndex: null,
        mapping: {},
        defaults: {},
        edits: {},
        validated: [],
        page: 1,
        pageSize: 50,
        duplicateMode: "skip",
        error: "",
        progress: 0,
        progressLabel: "",
        ocrWorker: null,
        abortRequested: false,
      };
    }

    function disposeImportSession(session) {
      if (session?.imageUrl) URL.revokeObjectURL(session.imageUrl);
      for (const url of session?.previewUrls || []) URL.revokeObjectURL(url);
      session?.ocrWorker?.terminate?.().catch?.(() => {});
    }

    async function renderImportCenter() {
      if (!state.importSession || state.importSession.type !== (state.importType || state.importSession.type)) {
        disposeImportSession(state.importSession);
        state.importSession = newImportSession(state.importType || "classes");
      }
      const session = state.importSession;
      state.importType = session.type;
      const [jobs, mappings, undoBatches] = await Promise.all([
        db.all("import_jobs"),
        db.all("import_mappings"),
        db.all("import_undo_batches"),
      ]);
      session.availableMappings = mappings.filter((row) => row.data_type === session.type);
      session.latestUndo = undoBatches
        .filter((row) => !row.undone_at)
        .sort((a, b) => String(b.committed_at || b.created_at).localeCompare(String(a.committed_at || a.created_at)))[0];
      const recentJobs = jobs
        .sort((a, b) => String(b.created_at).localeCompare(String(a.created_at)))
        .slice(0, 12);
      const typeOptions = Object.entries(engine.TYPES)
        .map(
          ([id, definition]) =>
            `<option value="${id}" ${session.type === id ? "selected" : ""}>${esc(definition.label)}</option>`,
        )
        .join("");
      setContent(
        pageHead(
          "Trung tâm nhập dữ liệu",
          "Một quy trình thống nhất: chọn nguồn, ánh xạ, kiểm tra, xem trước, snapshot, nhập giao dịch và hoàn tác.",
          `<button class="btn" id="importDownloadTemplate">Tải mẫu Excel</button>${session.latestUndo ? '<button class="btn danger" id="importUndoLatest">Hoàn tác lượt nhập gần nhất</button>' : ""}`,
        ) +
          `<div class="stepper"><span class="${session.status === "idle" ? "active" : ""}">1. Loại & nguồn</span><span class="${["reading", "mapped"].includes(session.status) ? "active" : ""}">2. Nhận diện</span><span class="${session.status === "ready" ? "active" : ""}">3. Ánh xạ & kiểm tra</span><span class="${session.status === "committing" ? "active" : ""}">4. Xác nhận nhập</span><span class="${session.status === "completed" ? "active" : ""}">5. Kết quả</span></div>
          <div class="import-grid"><aside class="card"><div class="card-head"><h2>Nguồn dữ liệu</h2></div><div class="card-body"><div class="field"><label class="required">Loại dữ liệu</label><select id="importType">${typeOptions}</select></div><div class="field mt"><label>Xử lý bản ghi trùng</label><select id="importDuplicateMode"><option value="skip" ${session.duplicateMode === "skip" ? "selected" : ""}>Bỏ qua bản ghi trùng</option><option value="fill" ${session.duplicateMode === "fill" ? "selected" : ""}>Chỉ điền trường đang trống</option><option value="update" ${session.duplicateMode === "update" ? "selected" : ""}>Cập nhật theo mã định danh</option><option value="create" ${session.duplicateMode === "create" ? "selected" : ""}>Tạo bản ghi mới có ghi dấu trùng</option></select></div><div class="drop-zone smart-import-drop mt" id="smartImportDrop"><div><strong>Kéo thả tệp vào đây</strong><br><small>XLSX, XLS, CSV, DOCX, PDF, PNG, JPG, JPEG, WEBP</small><br><button class="btn small mt" id="chooseImportFiles" type="button">Chọn tệp</button></div></div><div class="field mt"><label>Hoặc dán bảng từ Excel/Google Sheets</label><textarea id="importPasteText" placeholder="Dán bảng có hàng tiêu đề…" style="min-height:120px"></textarea></div><button class="btn primary mt" id="readImportPaste" style="width:100%">Đọc bảng đã dán</button>${session.files.length ? `<div class="notice mt"><strong>${session.files.length} tệp đã chọn</strong><br>${session.files.slice(0, 5).map((file) => esc(file.name)).join("<br>")}${session.files.length > 5 ? `<br>… và ${session.files.length - 5} tệp khác` : ""}</div>` : ""}${session.files.length && session.type === "documents" ? '<button class="btn mt" id="saveImportOriginals" style="width:100%">Lưu tệp gốc vào Hồ sơ</button>' : ""}</div></aside><section id="importWorkspace">${renderImportWorkspace(session)}</section></div>
          <div class="card mt"><div class="card-head"><h2>Lịch sử nhập gần đây</h2><span class="meta">${recentJobs.length} lượt</span></div><div class="card-body"><div class="table-wrap" style="max-height:300px"><table><thead><tr><th>Thời gian</th><th>Loại</th><th>Nguồn</th><th>Kết quả</th><th>Mới</th><th>Cập nhật</th><th>Bỏ qua</th><th>Lỗi</th></tr></thead><tbody>${recentJobs
            .map(
              (job) =>
                `<tr><td>${fmtDateTime(job.created_at)}</td><td>${esc(engine.TYPES[job.data_type]?.label || job.data_type)}</td><td class="wrap">${esc(job.source_name || "Dữ liệu dán")}</td><td>${job.status === "completed" ? '<span class="badge green">Hoàn tất</span>' : job.status === "failed" ? '<span class="badge red">Thất bại</span>' : '<span class="badge yellow">Đang xử lý</span>'}</td><td>${Number(job.inserted || 0)}</td><td>${Number(job.updated || 0)}</td><td>${Number(job.skipped || 0)}</td><td>${Number(job.errors || 0)}</td></tr>`,
            )
            .join("") || '<tr><td colspan="8" class="empty">Chưa có lịch sử nhập.</td></tr>'}</tbody></table></div></div></div>`,
      );
      bindImportCenter(session);
    }

    function bindImportCenter(session) {
      if (!$("#importContextFields")) {
        const contextFields = document.createElement("div");
        contextFields.id = "importContextFields";
        contextFields.className = "form-grid mt";
        contextFields.innerHTML = `<div class="field"><label>Năm học áp dụng</label><select id="importContextYear">${state.cache.years.map((row) => `<option value="${row.id}" ${row.id === state.yearId ? "selected" : ""}>${esc(row.name)}</option>`).join("")}</select></div><div class="field"><label>Học kỳ</label><select id="importContextSemester"><option value="all">Cả năm/học kỳ</option>${state.cache.sems.filter((row) => row.school_year_id === state.yearId).map((row) => `<option value="${row.id}" ${row.id === state.semesterId ? "selected" : ""}>${esc(row.name)}</option>`).join("")}</select></div><div class="field"><label>Tuần</label><select id="importContextWeek"><option value="">Không cố định</option>${state.cache.weeks.filter((row) => row.school_year_id === state.yearId).map((row) => `<option value="${row.id}" ${row.id === state.weekId ? "selected" : ""}>${esc(row.name)}</option>`).join("")}</select></div><div class="field"><label>Cơ sở áp dụng</label><select id="importContextCampus"><option value="all">Toàn trường/đọc từ tệp</option>${state.cache.campuses.map((row) => `<option value="${row.id}" ${row.id === state.campusId ? "selected" : ""}>${esc(row.name)}</option>`).join("")}</select></div>`;
        $("#importDuplicateMode").closest(".field").after(contextFields);
      }
      const refreshImportContext = async () => {
        await loadContext();
        if (session.tables.length) await prepareImport(session);
        renderImportCenter();
      };
      $("#importContextYear").onchange = async (event) => {
        state.yearId = event.target.value;
        state.semesterId = "all";
        state.weekId = state.cache.weeks.find((row) => row.school_year_id === state.yearId)?.id || "";
        await refreshImportContext();
      };
      $("#importContextSemester").onchange = async (event) => {
        state.semesterId = event.target.value;
        await refreshImportContext();
      };
      $("#importContextWeek").onchange = async (event) => {
        state.weekId = event.target.value;
        await refreshImportContext();
      };
      $("#importContextCampus").onchange = async (event) => {
        state.campusId = event.target.value;
        await refreshImportContext();
      };
      $("#importType").onchange = (event) => {
        disposeImportSession(session);
        state.importType = event.target.value;
        state.importSession = newImportSession(state.importType);
        renderImportCenter();
      };
      $("#importDuplicateMode").onchange = (event) => {
        session.duplicateMode = event.target.value;
      };
      $("#importDownloadTemplate").onclick = () => downloadImportTemplate(session.type, false);
      if ($("#importUndoLatest"))
        $("#importUndoLatest").onclick = () => undoLatestImport(session.latestUndo);
      $("#chooseImportFiles").onclick = () => $("#smartImportFile").click();
      $("#smartImportFile").onchange = async (event) => {
        const files = [...event.target.files];
        event.target.value = "";
        if (files.length) await handleImportFiles(files, session);
      };
      const drop = $("#smartImportDrop");
      ["dragenter", "dragover"].forEach((name) =>
        drop.addEventListener(name, (event) => {
          event.preventDefault();
          drop.classList.add("dragging");
        }),
      );
      ["dragleave", "drop"].forEach((name) =>
        drop.addEventListener(name, (event) => {
          event.preventDefault();
          drop.classList.remove("dragging");
        }),
      );
      drop.ondrop = async (event) => {
        const files = [...event.dataTransfer.files];
        if (files.length) await handleImportFiles(files, session);
      };
      $("#readImportPaste").onclick = async () => {
        const value = $("#importPasteText").value;
        if (!value.trim()) return toast("Hãy dán bảng dữ liệu.", "bad");
        session.sourceName = "Dữ liệu dán";
        session.sourceChecksum = await engine.checksum(value);
        setImportTables(session, [
          { name: "Bảng đã dán", rows: engine.parseDelimited(value) },
        ]);
        await prepareImport(session);
        renderImportCenter();
      };
      if (session.files.length && !$("#importFileManager")) {
        session.fileMeta ||= session.files.map((file) => ({
          name: file.name.replace(/\.[^.]+$/, ""),
          category: "Hồ sơ",
          related_module: "",
          related_record_id: "",
        }));
        session.previewUrls ||= session.files.map((file) =>
          /^image\//.test(file.type) ? URL.createObjectURL(file) : "",
        );
        const manager = document.createElement("details");
        manager.id = "importFileManager";
        manager.className = "mt";
        manager.innerHTML = `<summary><strong>Đổi tên, phân loại và gắn ${session.files.length} tệp gốc</strong></summary><div class="table-wrap mt" style="max-height:300px"><table><thead><tr><th>Xem</th><th>Tên hiển thị</th><th>Phân loại</th><th>Gắn với</th><th>ID bản ghi</th></tr></thead><tbody>${session.files.map((file, index) => `<tr><td>${session.previewUrls[index] ? `<img src="${esc(session.previewUrls[index])}" alt="" style="width:52px;height:42px;object-fit:cover;border-radius:5px">` : `<code>${esc(file.name.split(".").pop().toUpperCase())}</code>`}</td><td><input data-import-file-name="${index}" value="${esc(session.fileMeta[index].name)}" maxlength="150"></td><td><select data-import-file-category="${index}">${["Hồ sơ", "Kế hoạch", "Hoạt động", "Vi phạm", "Khen thưởng", "Báo cáo", "Minh chứng"].map((value) => `<option value="${value}" ${session.fileMeta[index].category === value ? "selected" : ""}>${value}</option>`).join("")}</select></td><td><select data-import-file-module="${index}"><option value="">Không gắn</option>${["plans", "tasks", "activities", "student_incidents", "commendations", "generated_reports"].map((value) => `<option value="${value}" ${session.fileMeta[index].related_module === value ? "selected" : ""}>${value}</option>`).join("")}</select></td><td><input data-import-file-related="${index}" value="${esc(session.fileMeta[index].related_record_id)}" placeholder="Tùy chọn"></td></tr>`).join("")}</tbody></table></div></details>`;
        $("#smartImportDrop").after(manager);
      }
      $$('[data-import-file-name]').forEach((input) => {
        input.oninput = () => (session.fileMeta[Number(input.dataset.importFileName)].name = input.value);
      });
      $$('[data-import-file-category]').forEach((select) => {
        select.onchange = () => (session.fileMeta[Number(select.dataset.importFileCategory)].category = select.value);
      });
      $$('[data-import-file-module]').forEach((select) => {
        select.onchange = () => (session.fileMeta[Number(select.dataset.importFileModule)].related_module = select.value);
      });
      $$('[data-import-file-related]').forEach((input) => {
        input.oninput = () => (session.fileMeta[Number(input.dataset.importFileRelated)].related_record_id = input.value.trim());
      });
      if (session.files.length && !$("#saveImportOriginals")) {
        const button = document.createElement("button");
        button.id = "saveImportOriginals";
        button.className = "btn mt";
        button.style.width = "100%";
        button.type = "button";
        button.textContent = "Lưu tệp gốc vào Hồ sơ";
        $("#smartImportDrop").after(button);
      }
      if ($("#saveImportOriginals"))
        $("#saveImportOriginals").onclick = async () => {
          try {
            await saveDocumentFiles(session.files, session.fileMeta || []);
            toast(`Đã lưu ${session.files.length} tệp gốc vào Hồ sơ.`);
          } catch (error) {
            toast("Không thể lưu tệp gốc: " + error.message, "bad");
          }
        };
      bindImportWorkspace(session);
    }

    function findBestImportTable(tables, type) {
      const schema = engine.TYPES[type];
      if (!schema || !tables.length) return { tableIndex: 0, headerIndex: 0 };
      let best = { tableIndex: 0, headerIndex: 0, score: -1 };
      tables.forEach((table, tableIndex) => {
        const headerIndex = engine.detectHeaderRow(table.rows, type),
          mapping = engine.autoMap(table.rows[headerIndex] || [], type),
          mappedFields = Object.keys(mapping),
          requiredFields = schema.fields
            .filter(([, , required]) => required)
            .map(([name]) => name),
          requiredMapped = requiredFields.filter((name) =>
            Number.isInteger(mapping[name]),
          ).length,
          dataRows = Math.max(0, table.rows.length - headerIndex - 1),
          score = requiredMapped * 100 + mappedFields.length * 10 + Math.min(dataRows, 9);
        if (score > best.score)
          best = { tableIndex, headerIndex, score };
      });
      return best;
    }

    function setImportTables(session, tables) {
      session.tables = tables.filter(
        (table) => Array.isArray(table.rows) && table.rows.length,
      );
      const selected = findBestImportTable(session.tables, session.type);
      session.tableIndex = selected.tableIndex;
      session.headerIndex = selected.headerIndex;
      session.mapping = {};
      session.defaults = {};
      session.edits = {};
      session.page = 1;
      session.status = session.tables.length ? "mapped" : "error";
      session.error = session.tables.length
        ? ""
        : "Không tìm thấy bảng hoặc nội dung có thể ánh xạ.";
    }

    async function verifyImportFiles(files) {
      const maxMb = Math.max(1, Number(await setting("max_file_mb")) || 25);
      const allowed = new Set([
        "xlsx",
        "xls",
        "csv",
        "tsv",
        "docx",
        "pdf",
        "png",
        "jpg",
        "jpeg",
        "webp",
      ]);
      for (const file of files) {
        const extension = String(file.name || "").split(".").pop().toLowerCase();
        if (!allowed.has(extension))
          throw new Error(`Tệp ${file.name} không thuộc định dạng được hỗ trợ.`);
        if (file.size > maxMb * 1024 * 1024)
          throw new Error(`Tệp ${file.name} vượt giới hạn ${maxMb} MB.`);
        const bytes = new Uint8Array(await file.slice(0, 16).arrayBuffer());
        const ascii = String.fromCharCode(...bytes);
        const zip = bytes[0] === 0x50 && bytes[1] === 0x4b;
        const compound =
          bytes[0] === 0xd0 &&
          bytes[1] === 0xcf &&
          bytes[2] === 0x11 &&
          bytes[3] === 0xe0;
        const valid =
          (["xlsx", "docx"].includes(extension) && zip) ||
          (extension === "xls" && compound) ||
          (extension === "pdf" && ascii.startsWith("%PDF-")) ||
          (extension === "png" &&
            bytes[0] === 0x89 &&
            bytes[1] === 0x50 &&
            bytes[2] === 0x4e &&
            bytes[3] === 0x47) ||
          (["jpg", "jpeg"].includes(extension) &&
            bytes[0] === 0xff &&
            bytes[1] === 0xd8 &&
            bytes[2] === 0xff) ||
          (extension === "webp" &&
            ascii.slice(0, 4) === "RIFF" &&
            ascii.slice(8, 12) === "WEBP") ||
          (["csv", "tsv"].includes(extension) && !bytes.includes(0));
        if (!valid)
          throw new Error(
            `Nội dung thật của ${file.name} không khớp phần mở rộng; tệp chưa được đọc.`,
          );
      }
    }

    async function handleImportFiles(files, session) {
      session.files = files;
      session.fileMeta = files.map((file) => ({
        name: file.name.replace(/\.[^.]+$/, ""),
        category: "Hồ sơ",
        related_module: "",
        related_record_id: "",
      }));
      session.status = "reading";
      session.progress = 5;
      session.progressLabel = "Đang kiểm tra tệp…";
      session.error = "";
      renderImportCenter();
      try {
        await verifyImportFiles(files);
        const file = files[0];
        const extension = file.name.split(".").pop().toLowerCase();
        const buffer = await file.arrayBuffer();
        session.sourceName = files.map((item) => item.name).join(", ");
        session.sourceChecksum = await engine.checksum(buffer);
        session.progress = 20;
        let tables = [];
        if (["csv", "tsv"].includes(extension)) {
          const content = new TextDecoder("utf-8").decode(buffer);
          tables = [{ name: file.name, rows: engine.parseDelimited(content, extension === "tsv" ? "\t" : undefined) }];
        } else if (["xlsx", "xls"].includes(extension)) {
          session.progressLabel = "Đang tải bộ đọc Excel cục bộ…";
          const XLSX = await loadXlsx();
          const workbook = XLSX.read(buffer, {
            type: "array",
            cellDates: false,
            raw: false,
            dense: true,
          });
          tables = workbook.SheetNames.map((name) => ({
            name,
            rows: XLSX.utils.sheet_to_json(workbook.Sheets[name], {
              header: 1,
              raw: false,
              defval: "",
              blankrows: false,
            }),
          }));
        } else if (extension === "docx") {
          session.progressLabel = "Đang đọc Word cục bộ…";
          const mammoth = await loadMammoth();
          const [htmlResult, textResult] = await Promise.all([
            mammoth.convertToHtml({ arrayBuffer: buffer }),
            mammoth.extractRawText({ arrayBuffer: buffer }),
          ]);
          const documentNode = new DOMParser().parseFromString(htmlResult.value, "text/html");
          tables = [...documentNode.querySelectorAll("table")].map((table, index) => ({
            name: `Bảng Word ${index + 1}`,
            rows: [...table.querySelectorAll("tr")].map((row) =>
              [...row.querySelectorAll("th,td")].map((cell) => cell.textContent.trim()),
            ),
          }));
          if (!tables.length) {
            const raw = textResult.value.trim();
            session.extractedText = raw;
            tables = [
              {
                name: "Nội dung Word",
                rows: engine.parseDelimited(raw.includes("\t") ? raw : `NoiDung\n${raw.replace(/\r?\n/g, " ")}`),
              },
            ];
          }
          session.extractionWarnings = [
            ...(htmlResult.messages || []),
            ...(textResult.messages || []),
          ].map((message) => message.message || String(message));
        } else if (extension === "pdf") {
          session.progressLabel = "Đang đọc PDF cục bộ…";
          const pdfjs = await loadPdfJs();
          const pdf = await pdfjs.getDocument({ data: new Uint8Array(buffer) }).promise;
          for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
            const page = await pdf.getPage(pageNumber);
            const content = await page.getTextContent();
            const groups = new Map();
            for (const item of content.items || []) {
              if (!item.str?.trim()) continue;
              const y = Math.round(item.transform?.[5] || 0);
              if (!groups.has(y)) groups.set(y, []);
              groups.get(y).push({ x: item.transform?.[4] || 0, text: item.str.trim() });
            }
            const rows = [...groups.entries()]
              .sort((a, b) => b[0] - a[0])
              .map(([, items]) => {
                const sorted = items.sort((a, b) => a.x - b.x);
                const cells = [];
                let lastX = null;
                for (const item of sorted) {
                  if (lastX != null && item.x - lastX > 65) cells.push(item.text);
                  else if (cells.length) cells[cells.length - 1] += ` ${item.text}`;
                  else cells.push(item.text);
                  lastX = item.x + item.text.length * 5;
                }
                return cells;
              });
            if (!rows.length && !session.ocrImageBlob) {
              const viewport = page.getViewport({ scale: 1.6 });
              const canvas = document.createElement("canvas");
              canvas.width = Math.ceil(viewport.width);
              canvas.height = Math.ceil(viewport.height);
              await page.render({
                canvasContext: canvas.getContext("2d", { alpha: false }),
                viewport,
              }).promise;
              session.ocrImageBlob = await new Promise((resolve, reject) =>
                canvas.toBlob(
                  (blob) =>
                    blob
                      ? resolve(blob)
                      : reject(new Error("Không thể tạo ảnh xem trước PDF scan.")),
                  "image/png",
                ),
              );
              session.imageUrl = URL.createObjectURL(session.ocrImageBlob);
              session.scannedPdfPage = pageNumber;
            }
            tables.push({ name: `Trang PDF ${pageNumber}`, rows });
          }
        } else if (["png", "jpg", "jpeg", "webp"].includes(extension)) {
          session.imageUrl = URL.createObjectURL(file);
          session.status = "mapped";
          session.tables = [];
          session.progress = 100;
          session.progressLabel = "Ảnh đã sẵn sàng; chọn OCR để nhận dạng chữ.";
          renderImportCenter();
          return;
        } else {
          throw new Error("Loại tệp chưa được hỗ trợ để trích xuất. Tệp .doc cũ chỉ được lưu làm hồ sơ.");
        }
        if (session.ocrImageBlob && !tables.some((table) => table.rows.length)) {
          session.tables = [];
          session.status = "mapped";
          session.progress = 100;
          session.progressLabel = `PDF scan trang ${session.scannedPdfPage}; sẵn sàng OCR ngoại tuyến.`;
          renderImportCenter();
          return;
        }
        setImportTables(session, tables);
        await prepareImport(session);
      } catch (error) {
        session.status = "error";
        session.error = error.message;
      }
      renderImportCenter();
    }

    async function prepareImport(session) {
      const table = session.tables[session.tableIndex];
      if (!table?.rows?.length) {
        session.validated = [];
        session.status = "mapped";
        return;
      }
      session.headerIndex = engine.detectHeaderRow(
        table.rows,
        session.type,
        Number.isInteger(session.headerIndex) ? session.headerIndex : undefined,
      );
      let headers = table.rows[session.headerIndex] || [];
      if (session.type === "score_entries" && !table.matrixExpanded) {
        const baseMapping = engine.autoMap(headers, session.type);
        const availableCriteria = await db.all("criteria");
        const criterionColumns = headers
          .map((header, index) => ({
            index,
            criterion: availableCriteria.find(
              (row) => engine.codeText(row.code) === engine.codeText(header),
            ),
          }))
          .filter((item) => item.criterion);
        if (
          Number.isInteger(baseMapping.class_code) &&
          !Number.isInteger(baseMapping.criterion_code) &&
          criterionColumns.length
        ) {
          const expanded = [["MaLop", "MaTieuChi", "Tuan", "Diem", "GhiChu"]];
          const defaultWeek =
            state.cache.weeks.find((row) => row.id === state.weekId)?.number || "";
          for (const source of table.rows.slice(session.headerIndex + 1)) {
            for (const item of criterionColumns) {
              const score = source[item.index];
              if (score === "" || score == null) continue;
              expanded.push([
                source[baseMapping.class_code],
                item.criterion.code,
                Number.isInteger(baseMapping.week)
                  ? source[baseMapping.week]
                  : defaultWeek,
                score,
                "Nhập từ ma trận lớp × tiêu chí",
              ]);
            }
          }
          table.rows = expanded;
          table.name += " • ma trận đã chuyển đổi";
          table.matrixExpanded = true;
          session.matrixDetected = true;
          session.headerIndex = 0;
          session.mapping = {};
          headers = table.rows[0];
        }
      }
      const headerKeys = headers.map(engine.keyText).filter(Boolean);
      session.headerWarnings = [
        ...new Set(
          headerKeys.filter(
            (value, index) => headerKeys.indexOf(value) !== index,
          ),
        ),
      ];
      if (!Object.keys(session.mapping).length)
        session.mapping = engine.autoMap(headers, session.type);
      const mapped = engine
        .mapRows(table.rows, session.type, session.headerIndex, session.mapping, {
          ...session.defaults,
          school_year_id: state.yearId,
        })
        .map((item) => ({
          ...item,
          values: { ...item.values, ...(session.edits[item.sourceRow] || {}) },
        }));
      const schema = engine.TYPES[session.type];
      const [existing, campuses, classes, students, criteria, weeks, equipment, scoreSheets, tasks, plans] =
        await Promise.all([
          db.all(schema.store),
          db.all("campuses"),
          db.all("classes"),
          db.all("students"),
          db.all("criteria"),
          db.all("school_weeks"),
          db.all("equipment"),
          db.all("weekly_score_sheets"),
          db.all("tasks"),
          db.all("plans"),
        ]);
      session.refs = {
        existing,
        campuses,
        classes,
        students,
        criteria,
        weeks,
        equipment,
        scoreSheets,
        tasks,
        plans,
      };
      session.validated = engine.validateRows(session.type, mapped, {
        context: { schoolYearId: state.yearId },
        existing,
        campuses,
        classes,
        students,
      });
      if (session.type === "classes") {
        for (const item of session.validated) {
          const campus = campuses.find(
            (row) =>
              engine.codeText(row.code) === engine.codeText(item.values.campus_code),
          );
          if (
            campus &&
            classes.some(
              (row) =>
                row.school_year_id === state.yearId &&
                row.campus_id === campus.id &&
                normalizeText(row.class_name || "").replace(/\s+/g, " ") ===
                  normalizeText(item.values.class_name || "").replace(/\s+/g, " "),
            )
          )
            item.errors.push("Tên lớp đã tồn tại trong cùng cơ sở và năm học");
          item.valid = item.errors.length === 0;
        }
      }
      if (session.type === "score_entries") {
        for (const item of session.validated) {
          const klass = classes.find(
            (row) => engine.codeText(row.code || row.class_name) === engine.codeText(item.values.class_code),
          );
          const criterion = criteria.find(
            (row) => engine.codeText(row.code) === engine.codeText(item.values.criterion_code),
          );
          const week = weeks.find(
            (row) => row.school_year_id === state.yearId && Number(row.number) === Number(item.values.week),
          );
          const sheet = scoreSheets.find(
            (row) => row.class_id === klass?.id && row.week_id === week?.id,
          );
          if (!criterion) item.errors.push("Không tìm thấy tiêu chí");
          if (
            criterion &&
            Number.isFinite(Number(item.values.score)) &&
            ((criterion.min != null && Number(item.values.score) < Number(criterion.min)) ||
              (criterion.max != null && Number(item.values.score) > Number(criterion.max)))
          )
            item.errors.push(
              `Điểm ngoài giới hạn ${criterion.min ?? "−∞"} đến ${criterion.max ?? "+∞"}`,
            );
          if (!week) item.errors.push("Không tìm thấy tuần");
          if (!sheet) item.errors.push("Chưa có bảng điểm cho lớp và tuần");
          if (sheet?.status === "locked" || sheet?.locked_at)
            item.errors.push("Bảng điểm đã khóa");
          item.valid = item.errors.length === 0;
        }
      }
      if (session.type === "task_check_items") {
        for (const item of session.validated) {
          const task = tasks.find(
            (row) =>
              engine.codeText(row.code || row.title) ===
              engine.codeText(item.values.task_code),
          );
          if (!task) item.errors.push("Không tìm thấy công việc");
          item.valid = item.errors.length === 0;
        }
      }
      if (session.type === "tasks") {
        for (const item of session.validated) {
          if (
            item.values.plan_code &&
            !plans.some(
              (plan) =>
                engine.codeText(plan.code) ===
                engine.codeText(item.values.plan_code),
            )
          )
            item.errors.push("Không tìm thấy mã kế hoạch liên kết");
          const progress = Number(item.values.progress || 0);
          if (!Number.isFinite(progress) || progress < 0 || progress > 100)
            item.errors.push("Tiến độ phải từ 0 đến 100");
          if (
            item.values.start_date &&
            item.values.due_date &&
            item.values.due_date < item.values.start_date
          )
            item.errors.push("Hạn hoàn thành phải từ ngày bắt đầu trở đi");
          if (
            item.values.completion_date &&
            item.values.start_date &&
            item.values.completion_date < item.values.start_date
          )
            item.errors.push("Ngày hoàn thành thực tế không được trước ngày bắt đầu");
          if (
            item.values.status === "done" &&
            !String(item.values.result || "").trim()
          )
            item.errors.push(
              "Công việc hoàn thành phải có kết quả thực hiện",
            );
          item.valid = item.errors.length === 0;
        }
      }
      if (session.type === "plans") {
        for (const item of session.validated) {
          const progress = Number(item.values.progress || 0);
          if (!Number.isFinite(progress) || progress < 0 || progress > 100)
            item.errors.push("Tiến độ phải từ 0 đến 100");
          if (
            item.values.start_date &&
            item.values.end_date &&
            item.values.end_date < item.values.start_date
          )
            item.errors.push("Ngày kết thúc phải sau ngày bắt đầu");
          if (
            item.values.status === "finished" &&
            !String(item.values.result || "").trim()
          )
            item.errors.push("Kế hoạch kết thúc phải có kết quả thực hiện");
          item.valid = item.errors.length === 0;
        }
      }
      session.status = "ready";
      session.progress = 100;
      session.progressLabel = "Đã kiểm tra dữ liệu";
    }

    function renderImportWorkspace(session) {
      if (session.status === "idle")
        return '<div class="card"><div class="card-body"><div class="empty"><strong>Bắt đầu bằng cách chọn tệp hoặc dán bảng.</strong><br>Chưa có dữ liệu nào được ghi vào ứng dụng.</div></div></div>';
      if (session.status === "reading" || session.status === "committing")
        return `<div class="card"><div class="card-head"><h2>${session.status === "reading" ? "Đang đọc nguồn" : "Đang ghi giao dịch"}</h2></div><div class="card-body"><p>${esc(session.progressLabel || "Đang xử lý…")}</p><div class="progress-track"><span style="width:${Math.max(0, Math.min(100, session.progress))}%"></span></div><button class="btn danger mt" id="cancelImportWork">Hủy tác vụ</button></div></div>`;
      if (session.status === "error")
        return `<div class="card"><div class="card-body"><div class="notice danger"><strong>Không thể đọc dữ liệu.</strong><br>${esc(session.error)}</div></div></div>`;
      if (session.status === "completed") {
        const result = session.result || {};
        return `<div class="card"><div class="card-head"><h2>Nhập dữ liệu hoàn tất</h2><span class="badge green">Đã ghi an toàn</span></div><div class="card-body"><div class="metric-row"><div class="metric"><strong>${Number(result.inserted || 0)}</strong><span>Thêm mới</span></div><div class="metric"><strong>${Number(result.updated || 0)}</strong><span>Cập nhật</span></div><div class="metric"><strong>${Number(result.skipped || 0)}</strong><span>Bỏ qua</span></div><div class="metric"><strong>${Number(result.errors || 0)}</strong><span>Lỗi</span></div><div class="metric"><strong>${session.validated.length}</strong><span>Tổng dòng</span></div></div><div class="notice mt">Snapshot trước nhập, checksum nguồn và nhật ký kết quả đã được lưu. Danh sách/bộ lọc đã được làm mới.</div><div class="toolbar mt"><button class="btn" id="downloadImportResult">Tải báo cáo kết quả</button><button class="btn danger" id="undoCompletedImport" ${session.latestUndo ? "" : "disabled"}>Hoàn tác lượt nhập</button><button class="btn primary" id="startAnotherImport">Nhập lượt khác</button></div></div></div>`;
      }
      if (session.imageUrl && !session.tables.length)
        return `<div class="card"><div class="card-head"><h2>Nhận dạng hình ảnh</h2></div><div class="card-body"><img class="rotate-preview" id="importImagePreview" src="${esc(session.imageUrl)}" alt="Ảnh chờ OCR"><div class="toolbar mt"><button class="btn" id="rotateImportImage">Xoay ảnh</button><button class="btn primary" id="runImportOcr">OCR tiếng Việt ngoại tuyến</button></div><div class="notice warn mt">OCR có thể sai. Nội dung chỉ chuyển sang bước ánh xạ sau khi người dùng kiểm tra.</div></div></div>`;
      const table = session.tables[session.tableIndex];
      if (!table)
        return '<div class="card"><div class="card-body"><div class="empty">Không có bảng để xử lý.</div></div></div>';
      const schema = engine.TYPES[session.type];
      const headers = table.rows[session.headerIndex] || [];
      const validCount = session.validated.filter((item) => item.valid).length;
      const errorCount = session.validated.length - validCount;
      const duplicateCount = session.validated.filter((item) => item.exists).length;
      const warningCount = session.validated.filter((item) => item.warnings.length).length;
      const model = pageSlice(session.validated, session.page, session.pageSize);
      session.page = model.page;
      const mappedFields = schema.fields.filter(([name]) => Number.isInteger(session.mapping[name]));
      const previewFields = mappedFields.length ? mappedFields : schema.fields.slice(0, 6);
      return `<div class="card"><div class="card-head"><h2>Nhận diện và ánh xạ</h2><span class="meta">${esc(session.sourceName || "Nguồn dữ liệu")}</span></div><div class="card-body">${session.extractionWarnings?.length ? `<div class="notice warn">Tài liệu Word có ${session.extractionWarnings.length} cảnh báo bố cục; cần kiểm tra preview.</div>` : ""}<div class="form-grid"><div class="field"><label>Sheet/bảng/trang</label><select id="importTableSelect">${session.tables.map((item, index) => `<option value="${index}" ${session.tableIndex === index ? "selected" : ""}>${esc(item.name)}</option>`).join("")}</select></div><div class="field"><label>Hàng tiêu đề</label><input id="importHeaderRow" type="number" min="1" max="${Math.min(20, table.rows.length)}" value="${session.headerIndex + 1}"></div><div class="field"><label>Mẫu ánh xạ đã lưu</label><select id="importMappingTemplate"><option value="">— Chọn mẫu —</option>${(session.availableMappings || []).map((row) => `<option value="${row.id}">${esc(row.name)}</option>`).join("")}</select></div></div><div class="toolbar mt"><button class="btn small" id="autoImportMapping">Tự ánh xạ</button><button class="btn small" id="clearImportMapping">Xóa ánh xạ</button><button class="btn small" id="saveImportMapping">Lưu mẫu ánh xạ</button><button class="btn small" id="applyImportMapping">Dùng mẫu đã chọn</button></div><div class="mapping-grid mt"><div class="mapping-head">Trường dữ liệu trong ứng dụng</div><div class="mapping-head">Cột trong nguồn</div><div class="mapping-head">Giá trị mặc định nếu ô trống</div>${schema.fields.map(([name, label, required]) => `<label>${esc(label)}${required ? " *" : ""}</label><select data-import-map="${name}"><option value="">— Không ánh xạ —</option>${headers.map((header, index) => `<option value="${index}" ${session.mapping[name] === index ? "selected" : ""}>${esc(header || `Cột ${index + 1}`)}</option>`).join("")}</select><input data-import-default="${name}" value="${esc(session.defaults[name] || "")}" placeholder="Tùy chọn">`).join("")}</div></div></div>
        <div class="metric-row mt"><div class="metric"><strong>${session.validated.length.toLocaleString("vi-VN")}</strong><span>Tổng dòng</span></div><div class="metric"><strong>${validCount.toLocaleString("vi-VN")}</strong><span>Hợp lệ</span></div><div class="metric"><strong>${errorCount.toLocaleString("vi-VN")}</strong><span>Dòng lỗi</span></div><div class="metric"><strong>${duplicateCount.toLocaleString("vi-VN")}</strong><span>Đã tồn tại</span></div><div class="metric"><strong>${warningCount.toLocaleString("vi-VN")}</strong><span>Cảnh báo</span></div></div>
        <div class="card mt"><div class="card-head"><h2>Xem trước</h2><div><button class="btn small" id="downloadImportErrors" ${errorCount ? "" : "disabled"}>Tải danh sách lỗi</button></div></div><div class="card-body"><div class="notice ${errorCount ? "danger" : warningCount || duplicateCount ? "warn" : ""}">${errorCount ? "Phải sửa hoặc loại bỏ dữ liệu sai định dạng hoặc sai liên kết trước khi nhập." : warningCount ? "Có nội dung đang để trống hoặc cần lưu ý. Ứng dụng vẫn cho phép nhập và sẽ hỏi xác nhận thêm một lần." : "Dữ liệu chưa được ghi. Hãy kiểm tra và xác nhận cách xử lý trùng."}</div><div class="table-wrap" style="max-height:480px"><table><thead><tr><th>Dòng</th><th>Kết quả</th>${previewFields.map(([, label]) => `<th>${esc(label)}</th>`).join("")}</tr></thead><tbody>${model.rows.map((item) => `<tr class="${item.errors.length ? "row-error" : item.warnings.length ? "row-warning" : ""}"><td>${item.sourceRow}</td><td class="wrap">${item.errors.length ? `<strong>${esc(item.errors.join("; "))}</strong>` : item.warnings.length ? esc(item.warnings.join("; ")) : "Hợp lệ"}</td>${previewFields.map(([name]) => `<td><input class="cell-edit" data-import-row="${item.sourceRow}" data-import-field="${name}" value="${esc(item.values[name] ?? "")}"></td>`).join("")}</tr>`).join("") || `<tr><td colspan="${previewFields.length + 2}" class="empty">Không có dòng dữ liệu.</td></tr>`}</tbody></table></div>${paginationHtml("importPreview", model)}<div class="toolbar mt"><button class="btn" id="cancelImportSession">Làm lại</button><button class="btn primary" id="commitImportData" ${!session.validated.length || errorCount ? "disabled" : ""}>Xác nhận nhập ${validCount.toLocaleString("vi-VN")} dòng</button></div></div></div>`;
    }

    function bindImportWorkspace(session) {
      if (session.headerWarnings?.length && $("#importHeaderRow")) {
        const notice = document.createElement("div");
        notice.id = "duplicateHeaderWarning";
        notice.className = "notice warn mt";
        notice.textContent = `Phát hiện tiêu đề cột trùng: ${session.headerWarnings.join(", ")}. Hãy đổi hàng tiêu đề hoặc ánh xạ thủ công.`;
        $("#importHeaderRow").closest(".form-grid").after(notice);
      }
      if (session.matrixDetected && $("#importHeaderRow")) {
        const notice = document.createElement("div");
        notice.className = "notice mt";
        notice.textContent =
          "Đã nhận diện ma trận lớp × tiêu chí và chuyển thành các dòng điểm để kiểm tra trước khi nhập.";
        $("#importHeaderRow").closest(".form-grid").after(notice);
      }
      if ($("#downloadImportResult"))
        $("#downloadImportResult").onclick = () =>
          exportCsv(
            [
              {
                source: session.sourceName,
                checksum: session.sourceChecksum,
                total: session.validated.length,
                ...session.result,
                completed_at: now(),
              },
            ],
            [
              { name: "source", label: "Nguồn" },
              { name: "checksum", label: "Checksum" },
              { name: "total", label: "Tổng dòng" },
              { name: "inserted", label: "Thêm mới" },
              { name: "updated", label: "Cập nhật" },
              { name: "skipped", label: "Bỏ qua" },
              { name: "errors", label: "Lỗi" },
              { name: "completed_at", label: "Hoàn tất" },
            ],
            `ket-qua-nhap-${session.type}-${today()}`,
          );
      if ($("#undoCompletedImport"))
        $("#undoCompletedImport").onclick = () => undoLatestImport(session.latestUndo);
      if ($("#startAnotherImport"))
        $("#startAnotherImport").onclick = () => {
          disposeImportSession(session);
          state.importSession = newImportSession(session.type);
          renderImportCenter();
        };
      if ($("#cancelImportWork"))
        $("#cancelImportWork").onclick = async () => {
          session.abortRequested = true;
          if (session.ocrWorker) await session.ocrWorker.terminate().catch(() => {});
          session.status = "error";
          session.error = "Tác vụ đã được người dùng hủy; chưa ghi dữ liệu chính.";
          renderImportCenter();
        };
      if ($("#rotateImportImage")) {
        session.rotation ||= 0;
        $("#rotateImportImage").onclick = () => {
          session.rotation = (session.rotation + 90) % 360;
          $("#importImagePreview").style.transform = `rotate(${session.rotation}deg)`;
        };
      }
      if ($("#runImportOcr"))
        $("#runImportOcr").onclick = () => runImportOcr(session);
      if ($("#importTableSelect"))
        $("#importTableSelect").onchange = async (event) => {
          session.tableIndex = Number(event.target.value);
          session.headerIndex = null;
          session.mapping = {};
          session.defaults = {};
          session.edits = {};
          session.page = 1;
          await prepareImport(session);
          renderImportCenter();
        };
      if ($("#importHeaderRow"))
        $("#importHeaderRow").onchange = async (event) => {
          session.headerIndex = Math.max(0, Number(event.target.value) - 1);
          session.mapping = {};
          session.edits = {};
          session.page = 1;
          await prepareImport(session);
          renderImportCenter();
        };
      $$('[data-import-map]').forEach((select) => {
        select.onchange = async () => {
          if (select.value === "") delete session.mapping[select.dataset.importMap];
          else session.mapping[select.dataset.importMap] = Number(select.value);
          session.page = 1;
          await prepareImport(session);
          renderImportCenter();
        };
      });
      $$('[data-import-default]').forEach((input) => {
        input.onchange = async () => {
          const field = input.dataset.importDefault;
          if (input.value.trim()) session.defaults[field] = input.value.trim();
          else delete session.defaults[field];
          session.page = 1;
          await prepareImport(session);
          renderImportCenter();
        };
      });
      if ($("#autoImportMapping"))
        $("#autoImportMapping").onclick = async () => {
          const table = session.tables[session.tableIndex];
          session.mapping = engine.autoMap(table.rows[session.headerIndex] || [], session.type);
          await prepareImport(session);
          renderImportCenter();
        };
      if ($("#clearImportMapping"))
        $("#clearImportMapping").onclick = async () => {
          session.mapping = {};
          await prepareImport(session);
          renderImportCenter();
        };
      if ($("#saveImportMapping"))
        $("#saveImportMapping").onclick = () => saveImportMapping(session);
      if ($("#applyImportMapping"))
        $("#applyImportMapping").onclick = () => applySavedMapping(session);
      $$('[data-import-row]').forEach((input) => {
        input.onchange = async () => {
          const row = Number(input.dataset.importRow);
          session.edits[row] ||= {};
          session.edits[row][input.dataset.importField] = input.value;
          await prepareImport(session);
          renderImportCenter();
        };
      });
      if ($("#importPreviewPrev"))
        $("#importPreviewPrev").onclick = () => {
          session.page -= 1;
          renderImportCenter();
        };
      if ($("#importPreviewNext"))
        $("#importPreviewNext").onclick = () => {
          session.page += 1;
          renderImportCenter();
        };
      if ($("#downloadImportErrors"))
        $("#downloadImportErrors").onclick = () => {
          const errorRows = session.validated
            .filter((item) => item.errors.length)
            .map((item) => ({
              source_row: item.sourceRow,
              errors: item.errors.join("; "),
              warnings: item.warnings.join("; "),
            }));
          exportCsv(
            errorRows,
            [
              { name: "source_row", label: "Dòng nguồn" },
              { name: "errors", label: "Lỗi" },
              { name: "warnings", label: "Cảnh báo" },
            ],
            `loi-nhap-${session.type}-${today()}`,
          );
        };
      if ($("#cancelImportSession"))
        $("#cancelImportSession").onclick = () => {
          disposeImportSession(session);
          state.importSession = newImportSession(session.type);
          renderImportCenter();
        };
      if ($("#commitImportData"))
        $("#commitImportData").onclick = () => commitImport(session);
    }

    async function rotatedImageForOcr(file, degrees) {
      const rotation = ((Number(degrees) % 360) + 360) % 360;
      if (!rotation) return file;
      const bitmap = await createImageBitmap(file);
      const swap = rotation === 90 || rotation === 270;
      const canvas = document.createElement("canvas");
      canvas.width = swap ? bitmap.height : bitmap.width;
      canvas.height = swap ? bitmap.width : bitmap.height;
      const context = canvas.getContext("2d", { alpha: false });
      context.translate(canvas.width / 2, canvas.height / 2);
      context.rotate((rotation * Math.PI) / 180);
      context.drawImage(bitmap, -bitmap.width / 2, -bitmap.height / 2);
      bitmap.close?.();
      return new Promise((resolve, reject) =>
        canvas.toBlob(
          (blob) => (blob ? resolve(blob) : reject(new Error("Không thể xoay ảnh để OCR."))),
          "image/png",
        ),
      );
    }

    async function runImportOcr(session) {
      const imageFile =
        session.ocrImageBlob || session.files.find((file) => /^image\//.test(file.type));
      if (!imageFile) return toast("Không tìm thấy ảnh để OCR.", "bad");
      session.status = "reading";
      session.progress = 1;
      session.progressLabel = "Đang tải OCR tiếng Việt cục bộ…";
      session.abortRequested = false;
      renderImportCenter();
      try {
        const Tesseract = await loadTesseract();
        const worker = await Tesseract.createWorker(
          "vie",
          Tesseract.OEM?.LSTM_ONLY ?? 1,
          {
            workerPath: new URL("./assets/vendor/tesseract/worker.min.js", location.href).href,
            corePath: new URL(
              "./assets/vendor/tesseract/tesseract-core-simd-lstm.wasm.js",
              location.href,
            ).href,
            langPath: new URL("./assets/vendor/tesseract/lang-data", location.href).href,
            logger: (message) => {
              if (message.progress != null) session.progress = Math.round(message.progress * 100);
              session.progressLabel = message.status || "Đang OCR…";
              const bar = $("#importWorkspace .progress-track > span");
              if (bar) bar.style.width = `${session.progress}%`;
            },
          },
        );
        session.ocrWorker = worker;
        const ocrSource = await rotatedImageForOcr(imageFile, session.rotation || 0);
        const result = await worker.recognize(ocrSource);
        await worker.terminate();
        session.ocrWorker = null;
        if (session.abortRequested) throw new Error("OCR đã bị hủy.");
        const extracted = result.data?.text?.trim() || "";
        if (!extracted) throw new Error("OCR không nhận diện được nội dung.");
        session.extractedText = extracted;
        setImportTables(session, [
          {
            name: "Kết quả OCR",
            rows: engine.parseDelimited(
              extracted.includes("\t") ? extracted : `NoiDung\n${extracted.replace(/\r?\n/g, " ")}`,
            ),
          },
        ]);
        await prepareImport(session);
      } catch (error) {
        session.status = "error";
        session.error = `${error.message} Nội dung OCR phải được kiểm tra trước khi lưu.`;
      }
      renderImportCenter();
    }

    function saveImportMapping(session) {
      const table = session.tables[session.tableIndex];
      const headers = table.rows[session.headerIndex] || [];
      openModal(
        "Lưu mẫu ánh xạ",
        `<div class="field"><label class="required">Tên mẫu</label><input id="importMappingName" value="${esc(engine.TYPES[session.type].label)} – mẫu 1" required maxlength="100"></div>`,
        `<button class="btn" id="cancelImportMapping">Hủy</button><button class="btn primary" id="confirmImportMapping">Lưu mẫu</button>`,
      );
      $("#cancelImportMapping").onclick = closeModal;
      $("#confirmImportMapping").onclick = async () => {
        const name = $("#importMappingName").value.trim();
        if (!name) return toast("Hãy nhập tên mẫu.", "bad");
        const columns = Object.fromEntries(
          Object.entries(session.mapping).map(([field, index]) => [field, headers[index]]),
        );
        const existing = (await db.all("import_mappings")).find(
          (row) => row.data_type === session.type && normalizeText(row.name) === normalizeText(name),
        );
        await db.put("import_mappings", {
          ...(existing || {}),
          name,
          data_type: session.type,
          columns,
          defaults: { ...session.defaults },
          active: true,
        });
        closeModal();
        toast("Đã lưu mẫu ánh xạ.");
        renderImportCenter();
      };
    }

    async function applySavedMapping(session) {
      const id = $("#importMappingTemplate").value;
      const template = (session.availableMappings || []).find((row) => row.id === id);
      if (!template) return toast("Hãy chọn mẫu ánh xạ.", "bad");
      const headers = session.tables[session.tableIndex].rows[session.headerIndex] || [];
      session.mapping = Object.fromEntries(
        Object.entries(template.columns || {})
          .map(([field, header]) => [field, headers.findIndex((value) => engine.keyText(value) === engine.keyText(header))])
          .filter(([, index]) => index >= 0),
      );
      session.defaults = { ...(template.defaults || {}) };
      await prepareImport(session);
      renderImportCenter();
    }

    function businessKey(type, row) {
      const campusCode = row.campus_code || state.cache.campuses?.find((item) => item.id === row.campus_id)?.code || "";
      if (type === "campuses") return engine.codeText(row.code);
      if (type === "classes")
        return [row.school_year_id || state.yearId, engine.codeText(campusCode), engine.codeText(row.code || row.class_name)].join("|");
      if (type === "students")
        return [row.school_year_id || state.yearId, engine.codeText(row.student_code)].join("|");
      if (type === "student_incidents") return engine.codeText(row.incident_code);
      if (type === "task_check_items")
        return [
          row.school_year_id || state.yearId,
          row.task_id || engine.codeText(row.task_code),
          engine.keyText(row.label),
        ].join("|");
      if (type === "equipment_transactions") return engine.codeText(row.code);
      if (type === "score_entries")
        return [row.school_year_id || state.yearId, row.week_id || row.week, row.class_id || engine.codeText(row.class_code), row.criterion_id || engine.codeText(row.criterion_code)].join("|");
      return engine.TYPES[type].key(row, { schoolYearId: state.yearId });
    }

    function hydrateImportedRow(type, row, refs) {
      const output = { ...row };
      const campus = refs.campuses.find(
        (item) => engine.codeText(item.code) === engine.codeText(output.campus_code),
      );
      if (campus) output.campus_id = campus.id;
      else if (state.campusId !== "all") output.campus_id ||= state.campusId;
      const klass = refs.classes.find(
        (item) =>
          item.school_year_id === state.yearId &&
          engine.codeText(item.code || item.class_name) === engine.codeText(output.class_code || output.class_name),
      );
      if (klass) {
        output.class_id = klass.id;
        output.class_code = klass.code || klass.class_name;
        output.class_name ||= klass.class_name;
        output.campus_id ||= klass.campus_id;
        output.campus_code ||= refs.campuses.find((item) => item.id === klass.campus_id)?.code || "";
      }
      if (type === "student_incidents") {
        const student = refs.students.find(
          (item) => engine.codeText(item.student_code) === engine.codeText(output.student_code),
        );
        if (student) {
          output.student_id = student.id;
          output.student_name ||= student.full_name;
          output.class_id ||= student.class_id;
          output.class_code ||= student.class_code;
          output.campus_id ||= student.campus_id;
          output.campus_code ||= student.campus_code;
        }
      }
      if (type === "score_entries") {
        const criterion = refs.criteria.find(
          (item) => engine.codeText(item.code) === engine.codeText(output.criterion_code),
        );
        const week = refs.weeks.find(
          (item) => item.school_year_id === state.yearId && Number(item.number) === Number(output.week),
        );
        const sheet = refs.scoreSheets.find(
          (item) => item.class_id === klass?.id && item.week_id === week?.id,
        );
        output.class_id = klass?.id;
        output.criterion_id = criterion?.id;
        output.week_id = week?.id;
        output.sheet_id = sheet?.id;
      }
      if (type === "equipment_transactions") {
        const equipment = refs.equipment.find(
          (item) => engine.codeText(item.code) === engine.codeText(output.equipment_code),
        );
        output.equipment_id = equipment?.id || null;
      }
      if (type === "task_check_items") {
        const task = refs.tasks.find(
          (item) =>
            engine.codeText(item.code || item.title) ===
            engine.codeText(output.task_code),
        );
        output.task_id = task?.id || null;
        if (output.required !== "" && output.required != null)
          output.required = ["1", "true", "co", "có", "yes"].includes(
            engine.keyText(output.required),
          );
        if (output.done !== "" && output.done != null)
          output.done = ["1", "true", "co", "có", "yes", "done"].includes(
            engine.keyText(output.done),
          );
        if (output.order !== "" && output.order != null)
          output.order = Number(output.order);
      }
      if (type === "tasks") {
        const plan = (refs.plans || []).find(
          (item) =>
            engine.codeText(item.code) === engine.codeText(output.plan_code),
        );
        output.plan_id = plan?.id || output.plan_id || null;
        if (output.progress !== "" && output.progress != null)
          output.progress = Number(output.progress);
        if (output.status === "done") output.progress = 100;
      }
      if (type === "plans") {
        if (output.progress !== "" && output.progress != null)
          output.progress = Number(output.progress);
        if (output.status === "finished") output.progress = 100;
      }
      if (type !== "campuses") {
        output.school_year_id ||= state.yearId;
        output.academic_year_id ||= state.yearId;
        if (state.semesterId !== "all") output.semester_id ||= state.semesterId;
        if (type !== "score_entries" && state.weekId)
          output.week_id ||= state.weekId;
      }
      if (type === "classes" && output.status)
        output.active = !["inactive", "false", "0", "ngung"].includes(
          engine.keyText(output.status),
        );
      if (type === "equipment" && output.quantity !== "" && output.quantity != null)
        output.quantity = Number(output.quantity);
      return output;
    }

    async function commitImport(session, confirmedBlankFields = false) {
      const invalid = session.validated.filter((item) => !item.valid);
      if (invalid.length)
        return toast("Còn dòng lỗi; chưa ghi dữ liệu.", "bad");
      if (!session.validated.length) return toast("Không có dòng để nhập.", "bad");
      const rowsWithBlanks = session.validated.filter(
        (item) => item.blankFields?.length,
      );
      if (rowsWithBlanks.length && !confirmedBlankFields) {
        const blankCounts = new Map();
        for (const item of rowsWithBlanks)
          for (const field of item.blankFields || [])
            blankCounts.set(
              field.label,
              (blankCounts.get(field.label) || 0) + 1,
            );
        session.blankSummary = [...blankCounts.entries()].map(
          ([label, count]) => ({ label, count }),
        );
        openModal(
          "Xác nhận nhập dữ liệu còn trống",
          `<div class="notice warn"><strong>${rowsWithBlanks.length.toLocaleString("vi-VN")} dòng có nội dung đang để trống.</strong><br>Các ô trống sẽ được giữ nguyên; ứng dụng không từ chối lượt nhập.</div><div class="table-wrap mt" style="max-height:360px"><table><thead><tr><th>Trường đang trống</th><th>Số dòng</th></tr></thead><tbody>${session.blankSummary.map((item) => `<tr><td>${esc(item.label)}</td><td>${item.count.toLocaleString("vi-VN")}</td></tr>`).join("")}</tbody></table></div><p class="meta mt">Dữ liệu sai định dạng hoặc mã liên kết không tồn tại vẫn phải sửa để tránh ghép nhầm dữ liệu.</p>`,
          '<button class="btn" id="cancelBlankImport">Quay lại bổ sung</button><button class="btn primary" id="confirmBlankImport">Vẫn nhập dữ liệu</button>',
          true,
        );
        $("#cancelBlankImport").onclick = closeModal;
        $("#confirmBlankImport").onclick = () => {
          closeModal();
          commitImport(session, true);
        };
        return;
      }
      session.status = "committing";
      session.progress = 3;
      session.progressLabel = "Đang tạo snapshot bảo vệ…";
      session.abortRequested = false;
      renderImportCenter();
      const schema = engine.TYPES[session.type];
      let job = null;
      let snapshot = null;
      let mainWritten = false;
      try {
        snapshot = await createInternalSnapshot(
          `Trước nhập ${schema.label} – ${session.sourceName || "dữ liệu dán"}`,
          {
            tier: "protected",
            protectedSnapshot: true,
            reason: "before-smart-import",
            yearId: state.yearId,
          },
        );
        if (session.abortRequested) throw new DOMException("Đã hủy", "AbortError");
        job = await db.put(
          "import_jobs",
          {
            data_type: session.type,
            source_name: session.sourceName || "Dữ liệu dán",
            source_checksum: session.sourceChecksum,
            source_size: session.files.reduce((sum, file) => sum + file.size, 0),
            total_rows: session.validated.length,
            blank_rows: rowsWithBlanks.length,
            blank_fields: session.blankSummary || [],
            duplicate_mode: session.duplicateMode,
            status: "started",
            started_at: now(),
            user: "local-session",
          },
          { audit: false },
        );
        session.progress = 18;
        session.progressLabel = "Đang đưa dữ liệu hợp lệ vào vùng staging…";
        const staging = await db.put(
          "import_staging",
          {
            import_job_id: job.id,
            data_type: session.type,
            rows: session.validated.map((item) => ({
              source_row: item.sourceRow,
              values: item.values,
            })),
            status: "validated",
          },
          { audit: false, sync: false },
        );
        const existing = session.refs.existing;
        const byKey = new Map(existing.map((row) => [businessKey(session.type, row), row]));
        const writes = [];
        let inserted = 0;
        let updated = 0;
        let skipped = 0;
        for (const item of session.validated) {
          const incoming = hydrateImportedRow(session.type, item.values, session.refs);
          const existingRow = item.blankFields?.some((field) => field.required)
            ? null
            : byKey.get(businessKey(session.type, incoming));
          if (existingRow && session.duplicateMode === "skip") {
            skipped += 1;
            continue;
          }
          if (existingRow && session.duplicateMode === "fill") {
            const merged = { ...existingRow };
            for (const [key, value] of Object.entries(incoming))
              if ((merged[key] == null || merged[key] === "") && value !== "" && value != null)
                merged[key] = value;
            writes.push(merged);
            updated += 1;
          } else if (existingRow && session.duplicateMode === "update") {
            writes.push({ ...existingRow, ...incoming, id: existingRow.id, revision: existingRow.revision });
            updated += 1;
          } else {
            writes.push({
              ...incoming,
              id: uid(),
              ...(existingRow ? { duplicate_of: existingRow.id } : {}),
              import_job_id: job.id,
              source: "smart-import",
            });
            inserted += 1;
          }
        }
        if (session.abortRequested) throw new DOMException("Đã hủy", "AbortError");
        session.progress = 45;
        session.progressLabel = `Đang ghi ${writes.length.toLocaleString("vi-VN")} bản ghi trong một giao dịch…`;
        if (writes.length) {
          await db.bulkPut(schema.store, writes);
          if (session.type === "students")
            for (const student of writes) {
              const previous = existing.find(
                (row) => row.id === student.id,
              );
              await propagateStudentIdentity(student, previous);
            }
          mainWritten = true;
        }
        if (session.abortRequested) throw new DOMException("Đã hủy", "AbortError");
        session.progress = 88;
        session.progressLabel = "Đang hoàn tất nhật ký và khả năng hoàn tác…";
        await db.put(
          "import_undo_batches",
          {
            import_job_id: job.id,
            snapshot_id: snapshot.id,
            data_type: session.type,
            store: schema.store,
            inserted,
            updated,
            skipped,
            committed_at: now(),
          },
          { audit: false, sync: false },
        );
        await db.put(
          "import_jobs",
          {
            ...job,
            status: "completed",
            inserted,
            updated,
            skipped,
            errors: 0,
            completed_at: now(),
            duration_ms: Date.now() - Date.parse(job.started_at),
          },
          { audit: false },
        );
        await db.hardDelete("import_staging", staging.id, true);
        if (session.type === "documents" && session.files.length)
          await saveDocumentFiles(session.files, session.fileMeta || []);
        session.status = "completed";
        session.progress = 100;
        session.result = {
          inserted,
          updated,
          skipped,
          errors: 0,
          duration_ms: Date.now() - Date.parse(job.started_at),
          blank_rows: rowsWithBlanks.length,
        };
        session.progressLabel = "Hoàn tất";
        await loadContext();
        toast(
          `Đã nhập ${inserted} mới, cập nhật ${updated}, bỏ qua ${skipped}${rowsWithBlanks.length ? `; giữ trống ${rowsWithBlanks.length} dòng` : ""}.`,
        );
      } catch (error) {
        if (mainWritten && snapshot?.payload) {
          try {
            await db.replaceAll({ schema: snapshot.schema, data: snapshot.payload });
            mainWritten = false;
          } catch (rollbackError) {
            error.message += `; khôi phục snapshot thất bại: ${rollbackError.message}`;
          }
        }
        session.status = "error";
        session.error =
          error.name === "AbortError"
            ? "Đã hủy trước khi giao dịch ghi dữ liệu chính hoàn tất."
            : `Lượt nhập thất bại; giao dịch chính không tạo bản ghi một phần. ${error.message}`;
        if (job)
          await db
            .put(
              "import_jobs",
              {
                ...job,
                status: "failed",
                errors: session.validated.length,
                error_message: error.message,
                completed_at: now(),
              },
              { audit: false },
            )
            .catch(() => {});
      }
      renderImportCenter();
    }

    async function undoLatestImport(batch) {
      if (!batch) return;
      const snapshot = await db.get("internal_snapshots", batch.snapshot_id);
      if (!snapshot?.payload)
        return toast("Không tìm thấy snapshot của lượt nhập.", "bad");
      openModal(
        "Hoàn tác lượt nhập gần nhất",
        `<div class="notice danger"><strong>Khôi phục toàn bộ dữ liệu nghiệp vụ về trước lượt nhập.</strong><br>Lượt nhập: ${esc(engine.TYPES[batch.data_type]?.label || batch.data_type)} • ${fmtDateTime(batch.committed_at)}.</div><p>Một snapshot mới sẽ được tạo trước khi hoàn tác.</p>`,
        `<button class="btn" id="cancelUndoImport">Hủy</button><button class="btn danger" id="confirmUndoImport">Xác nhận hoàn tác</button>`,
      );
      $("#cancelUndoImport").onclick = closeModal;
      $("#confirmUndoImport").onclick = async () => {
        const button = $("#confirmUndoImport");
        button.disabled = true;
        try {
          await createInternalSnapshot("Trước hoàn tác lượt nhập", {
            tier: "protected",
            protectedSnapshot: true,
            reason: "before-import-undo",
          });
          await db.replaceAll({ schema: snapshot.schema, data: snapshot.payload });
          await db.put(
            "import_undo_batches",
            { ...batch, undone_at: now(), status: "undone" },
            { audit: false, sync: false },
          );
          closeModal();
          await loadContext();
          toast("Đã hoàn tác lượt nhập và giữ snapshot bảo vệ.");
          state.importSession = newImportSession(batch.data_type);
          renderImportCenter();
        } catch (error) {
          button.disabled = false;
          toast("Hoàn tác thất bại: " + error.message, "bad");
        }
      };
    }

    async function renderClassManagement(panel) {
      const [classes, students, campuses] = await Promise.all([
        db.all("classes"),
        db.all("students"),
        db.all("campuses"),
      ]);
      const query = normalizeText(state.classQuery);
      const rows = classes
        .filter((row) => row.school_year_id === state.yearId)
        .filter((row) => state.campusId === "all" || row.campus_id === state.campusId)
        .filter(
          (row) => state.classGradeFilter === "all" || String(row.grade) === state.classGradeFilter,
        )
        .filter(
          (row) =>
            state.classStatusFilter === "all" ||
            (state.classStatusFilter === "active" ? row.active !== false : row.active === false),
        )
        .filter((row) =>
          query
            ? normalizeText([row.code, row.class_name, row.teacher, campusName(row.campus_id)].join(" ")).includes(query)
            : true,
        )
        .sort((a, b) => String(a.class_name).localeCompare(String(b.class_name), "vi", { numeric: true }));
      const counts = students.reduce((map, row) => {
        if (row.status !== "inactive") map[row.class_id] = (map[row.class_id] || 0) + 1;
        return map;
      }, {});
      const model = pageSlice(rows, state.classPage, state.classPageSize);
      state.classPage = model.page;
      const selected = state.classSelected;
      panel.innerHTML = `<div class="metric-row"><div class="metric"><strong>${rows.length}</strong><span>Lớp theo bộ lọc</span></div><div class="metric"><strong>${rows.filter((row) => row.active !== false).length}</strong><span>Đang sử dụng</span></div><div class="metric"><strong>${rows.filter((row) => row.active === false).length}</strong><span>Ngừng sử dụng</span></div><div class="metric"><strong>${students.filter((row) => row.school_year_id === state.yearId && row.status !== "inactive").length.toLocaleString("vi-VN")}</strong><span>Học sinh hiện hành</span></div><div class="metric"><strong>${selected.size}</strong><span>Lớp đã chọn</span></div></div><div class="card mt"><div class="card-head"><h2>Lớp và giáo viên chủ nhiệm</h2><span class="meta">Phân trang ${state.classPageSize} dòng</span></div><div class="card-body"><div class="toolbar"><input class="grow" id="classCenterSearch" value="${esc(state.classQuery)}" placeholder="Tìm mã, tên lớp, giáo viên…"><select id="classCenterGrade"><option value="all">Tất cả khối</option>${[1, 2, 3, 4, 5, 6, 7, 8, 9].map((grade) => `<option value="${grade}" ${state.classGradeFilter === String(grade) ? "selected" : ""}>Khối ${grade}</option>`).join("")}</select><select id="classCenterStatus"><option value="all">Tất cả trạng thái</option><option value="active" ${state.classStatusFilter === "active" ? "selected" : ""}>Đang dùng</option><option value="inactive" ${state.classStatusFilter === "inactive" ? "selected" : ""}>Ngừng dùng</option></select><button class="btn small" id="classCenterExportCsv">CSV</button><button class="btn small" id="classCenterExportXlsx">Excel</button></div><div class="toolbar mt"><button class="btn primary" id="classCenterAdd">＋ Thêm lớp</button><button class="btn" id="classCenterImport">Trung tâm nhập dữ liệu</button><button class="btn" id="classCenterTemplate">Tải mẫu Excel</button></div>${selected.size ? `<div class="bulkbar"><strong>${selected.size} lớp đã chọn</strong><button class="btn small" id="classBulkMove">Chuyển cơ sở</button><button class="btn small" id="classBulkInactive">Ngừng sử dụng</button><button class="btn small" id="classClearSelection">Bỏ chọn</button></div>` : ""}<div class="notice">Chỉ xóa được lớp chưa có học sinh, thi đua, vi phạm, hoạt động hoặc báo cáo. Lớp đã phát sinh dữ liệu chỉ được ngừng sử dụng.</div><div class="table-wrap" style="max-height:500px"><table><thead><tr><th><input type="checkbox" id="classSelectPage" aria-label="Chọn trang lớp hiện tại"></th><th>Mã</th><th>Lớp</th><th>Khối</th><th>Cơ sở</th><th>GVCN</th><th>Học sinh</th><th>Trạng thái</th><th>Thao tác</th></tr></thead><tbody>${model.rows
        .map(
          (row) =>
            `<tr><td><input type="checkbox" data-class-select="${row.id}" ${selected.has(row.id) ? "checked" : ""}></td><td><code>${esc(row.code || "—")}</code></td><td><strong>${esc(row.class_name)}</strong></td><td>${esc(row.grade)}</td><td>${esc(campusName(row.campus_id))}</td><td class="wrap">${esc(row.teacher || "—")}</td><td>${Number(counts[row.id] || 0)}</td><td>${row.active !== false ? '<span class="badge green">Đang dùng</span>' : '<span class="badge">Ngừng dùng</span>'}</td><td><button class="link-btn" data-class-detail="${row.id}">Xem</button><button class="link-btn" data-class-edit="${row.id}">Sửa</button><button class="link-btn" data-class-clone="${row.id}">Nhân bản</button><button class="link-btn" data-class-move="${row.id}">Chuyển</button><button class="link-btn red" data-class-delete="${row.id}">${row.active !== false ? "Xóa/Ngừng" : "Xóa"}</button></td></tr>`,
        )
        .join("") || '<tr><td colspan="9" class="empty">Chưa có lớp theo bộ lọc.</td></tr>'}</tbody></table></div>${paginationHtml("classCenter", model)}</div></div>`;
      let timer;
      $("#classCenterSearch").oninput = (event) => {
        clearTimeout(timer);
        timer = setTimeout(() => {
          state.classQuery = event.target.value;
          state.classPage = 1;
          renderClassManagement(panel);
        }, 250);
      };
      $("#classCenterGrade").onchange = (event) => {
        state.classGradeFilter = event.target.value;
        state.classPage = 1;
        renderClassManagement(panel);
      };
      $("#classCenterStatus").onchange = (event) => {
        state.classStatusFilter = event.target.value;
        state.classPage = 1;
        renderClassManagement(panel);
      };
      $("#classCenterPrev").onclick = () => {
        state.classPage -= 1;
        renderClassManagement(panel);
      };
      $("#classCenterNext").onclick = () => {
        state.classPage += 1;
        renderClassManagement(panel);
      };
      $("#classSelectPage").onchange = (event) => {
        for (const row of model.rows)
          if (event.target.checked) selected.add(row.id);
          else selected.delete(row.id);
        renderClassManagement(panel);
      };
      $$('[data-class-select]').forEach((input) => {
        input.onchange = () => {
          if (input.checked) selected.add(input.dataset.classSelect);
          else selected.delete(input.dataset.classSelect);
          renderClassManagement(panel);
        };
      });
      $("#classCenterAdd").onclick = () => classForm();
      $("#classCenterImport").onclick = () => {
        state.importType = "classes";
        state.importSession = null;
        go("imports");
      };
      $("#classCenterTemplate").onclick = () => downloadImportTemplate("classes", true);
      const exportRows = rows.map((row) => ({
        code: row.code || "",
        class_name: row.class_name,
        grade: row.grade,
        campus_code: campuses.find((item) => item.id === row.campus_id)?.code || "",
        campus_name: campusName(row.campus_id),
        teacher: row.teacher || "",
        size: counts[row.id] || row.size || 0,
        status: row.active !== false ? "active" : "inactive",
        notes: row.notes || "",
      }));
      const columns = [
        { name: "code", label: "Mã lớp" },
        { name: "class_name", label: "Tên lớp" },
        { name: "grade", label: "Khối" },
        { name: "campus_code", label: "Mã cơ sở" },
        { name: "campus_name", label: "Tên cơ sở" },
        { name: "teacher", label: "Giáo viên chủ nhiệm" },
        { name: "size", label: "Sĩ số" },
        { name: "status", label: "Trạng thái" },
        { name: "notes", label: "Ghi chú" },
      ];
      $("#classCenterExportCsv").onclick = () => exportCsv(exportRows, columns, `danh-sach-lop-${today()}`);
      $("#classCenterExportXlsx").onclick = () =>
        exportXlsx(exportRows, columns, `danh-sach-lop-${today()}`, "DANH SÁCH LỚP").catch((error) => toast(error.message, "bad"));
      $$('[data-class-edit]').forEach(
        (button) => (button.onclick = () => classForm(button.dataset.classEdit)),
      );
      $$('[data-class-detail]').forEach(
        (button) =>
          (button.onclick = () => {
            const row = classes.find((item) => item.id === button.dataset.classDetail);
            if (!row) return;
            openModal(
              "Chi tiết lớp",
              `<div class="split"><span>Mã lớp</span><code>${esc(row.code || "—")}</code></div><div class="split mt"><span>Tên lớp</span><strong>${esc(row.class_name)}</strong></div><div class="split mt"><span>Khối</span><span>${esc(row.grade)}</span></div><div class="split mt"><span>Cơ sở</span><span>${esc(campusName(row.campus_id))}</span></div><div class="split mt"><span>Giáo viên chủ nhiệm</span><span>${esc(row.teacher || "—")}</span></div><div class="split mt"><span>Học sinh hiện hành</span><strong>${Number(counts[row.id] || 0)}</strong></div>`,
              `<button class="btn" id="closeClassDetail">Đóng</button>`,
            );
            $("#closeClassDetail").onclick = closeModal;
          }),
      );
      $$('[data-class-clone]').forEach(
        (button) =>
          (button.onclick = async () => {
            const row = classes.find((item) => item.id === button.dataset.classClone);
            if (!row) return;
            let suffix = 1;
            let cloneCode = `${row.code || row.class_name}-COPY`;
            let cloneName = `${row.class_name} – bản sao`;
            while (
              classes.some(
                (item) =>
                  item.school_year_id === row.school_year_id &&
                  (engine.codeText(item.code) === engine.codeText(cloneCode) ||
                    (item.campus_id === row.campus_id &&
                      normalizeText(item.class_name) === normalizeText(cloneName))),
              )
            ) {
              suffix += 1;
              cloneCode = `${row.code || row.class_name}-COPY-${suffix}`;
              cloneName = `${row.class_name} – bản sao ${suffix}`;
            }
            const clone = await db.put("classes", {
              ...row,
              id: uid(),
              code: cloneCode,
              class_name: cloneName,
              teacher: "",
              size: 0,
              revision: undefined,
              created_at: undefined,
              updated_at: undefined,
            });
            toast("Đã nhân bản lớp; chưa sao chép học sinh hoặc dữ liệu phát sinh.");
            classForm(clone.id);
          }),
      );
      $$('[data-class-move]').forEach(
        (button) =>
          (button.onclick = () => {
            selected.clear();
            selected.add(button.dataset.classMove);
            moveSelectedClasses(panel, campuses);
          }),
      );
      $$('[data-class-delete]').forEach(
        (button) =>
          (button.onclick = () => deleteOrDeactivateClass(button.dataset.classDelete, panel)),
      );
      if ($("#classClearSelection"))
        $("#classClearSelection").onclick = () => {
          selected.clear();
          renderClassManagement(panel);
        };
      if ($("#classBulkMove"))
        $("#classBulkMove").onclick = () => moveSelectedClasses(panel, campuses);
      if ($("#classBulkInactive"))
        $("#classBulkInactive").onclick = () => deactivateSelectedClasses(panel);
    }

    async function moveSelectedClasses(panel, campuses) {
      const count = state.classSelected.size;
      if (!count) return;
      openModal(
        `Chuyển ${count} lớp sang cơ sở khác`,
        `<div class="notice warn">Học sinh hiện hành sẽ chuyển cùng lớp. Vi phạm, thi đua và báo cáo lịch sử giữ nguyên cơ sở đã ghi nhận.</div><div class="field mt"><label class="required">Cơ sở đích</label><select id="classTargetCampus"><option value="">— Chọn cơ sở —</option>${campuses.map((row) => `<option value="${row.id}">${esc(row.name)} (${esc(row.code || "—")})</option>`).join("")}</select></div>`,
        `<button class="btn" id="cancelClassMove">Hủy</button><button class="btn primary" id="confirmClassMove">Xác nhận chuyển</button>`,
      );
      $("#cancelClassMove").onclick = closeModal;
      $("#confirmClassMove").onclick = async () => {
        const campus = campuses.find((row) => row.id === $("#classTargetCampus").value);
        if (!campus) return toast("Hãy chọn cơ sở đích.", "bad");
        const [classes, students] = await Promise.all([
          db.all("classes"),
          db.all("students"),
        ]);
        const selectedClasses = classes.filter((row) => state.classSelected.has(row.id));
        await createInternalSnapshot("Trước chuyển cơ sở hàng loạt", {
          tier: "protected",
          protectedSnapshot: true,
          reason: "before-bulk-class-campus",
          yearId: state.yearId,
        });
        await db.bulkPut(
          "classes",
          selectedClasses.map((row) => ({ ...row, campus_id: campus.id })),
        );
        const movedStudents = students.filter((row) => state.classSelected.has(row.class_id));
        if (movedStudents.length)
          await db.bulkPut(
            "students",
            movedStudents.map((row) => ({
              ...row,
              campus_id: campus.id,
              campus_code: campus.code || "",
            })),
          );
        state.classSelected.clear();
        closeModal();
        toast(`Đã chuyển ${selectedClasses.length} lớp và ${movedStudents.length} học sinh hiện hành.`);
        renderClassManagement(panel);
      };
    }

    async function deactivateSelectedClasses(panel) {
      const classes = (await db.all("classes")).filter((row) => state.classSelected.has(row.id));
      if (!classes.length) return;
      openModal(
        "Ngừng sử dụng hàng loạt",
        `<div class="notice warn">Sẽ ngừng sử dụng ${classes.length} lớp. Dữ liệu lịch sử và học sinh không bị xóa.</div>`,
        `<button class="btn" id="cancelClassDeactivate">Hủy</button><button class="btn primary" id="confirmClassDeactivate">Xác nhận</button>`,
      );
      $("#cancelClassDeactivate").onclick = closeModal;
      $("#confirmClassDeactivate").onclick = async () => {
        await createInternalSnapshot("Trước ngừng sử dụng lớp hàng loạt", {
          tier: "protected",
          protectedSnapshot: true,
          reason: "before-bulk-class-inactive",
          yearId: state.yearId,
        });
        await db.bulkPut(
          "classes",
          classes.map((row) => ({ ...row, active: false })),
        );
        state.classSelected.clear();
        closeModal();
        toast(`Đã ngừng sử dụng ${classes.length} lớp.`);
        renderClassManagement(panel);
      };
    }

    async function deleteOrDeactivateClass(id, panel) {
      const klass = await db.get("classes", id);
      if (!klass) return;
      const stores = [
        ["students", "Học sinh", (row) => row.class_id === id],
        ["weekly_score_sheets", "Bảng thi đua", (row) => row.class_id === id],
        ["student_incidents", "Vi phạm", (row) => row.class_id === id],
        ["activity_classes", "Hoạt động", (row) => row.class_id === id],
        ["generated_reports", "Báo cáo", (row) => row.class_id === id],
      ];
      const usage = [];
      for (const [store, label, predicate] of stores) {
        const count = (await db.all(store)).filter(predicate).length;
        if (count) usage.push({ label, count });
      }
      openModal(
        usage.length ? "Lớp đã có dữ liệu liên quan" : "Xóa lớp chưa phát sinh dữ liệu",
        usage.length
          ? `<div class="notice warn"><strong>Không thể xóa ${esc(klass.class_name)}.</strong><br>${usage.map((item) => `${esc(item.label)}: ${item.count}`).join("<br>")}</div><p>Chỉ có thể ngừng sử dụng; tên lớp trong lịch sử được giữ nguyên.</p>`
          : `<div class="notice danger">Lớp ${esc(klass.class_name)} chưa có dữ liệu liên quan và có thể xóa mềm.</div>`,
        `<button class="btn" id="cancelClassDelete">Hủy</button><button class="btn ${usage.length ? "primary" : "danger"}" id="confirmClassDelete">${usage.length ? "Ngừng sử dụng" : "Xóa lớp"}</button>`,
      );
      $("#cancelClassDelete").onclick = closeModal;
      $("#confirmClassDelete").onclick = async () => {
        if (usage.length) await db.put("classes", { ...klass, active: false });
        else await db.remove("classes", id);
        state.classSelected.delete(id);
        closeModal();
        toast(usage.length ? "Đã ngừng sử dụng lớp." : "Đã xóa lớp chưa phát sinh dữ liệu.");
        renderClassManagement(panel);
      };
    }

    async function downloadImportTemplate(type, illustrated = false) {
      const schema = engine.TYPES[type];
      if (!schema) throw new Error("Chưa có mẫu cho loại dữ liệu này.");
      try {
        const XLSX = await loadXlsx();
        const headers = TEMPLATE_COLUMNS[type] || schema.fields.map(([name]) => name);
        const example = TEMPLATE_EXAMPLES[type] || headers.map(() => "");
        const rows = illustrated ? [headers, example] : [headers];
        const dataSheet = XLSX.utils.aoa_to_sheet(rows);
        dataSheet["!cols"] = headers.map((header, index) => ({
          wch: Math.min(32, Math.max(14, header.length + 2, String(example[index] || "").length + 2)),
        }));
        if (headers.length)
          dataSheet["!autofilter"] = {
            ref: `A1:${XLSX.utils.encode_col(headers.length - 1)}${rows.length}`,
          };
        const book = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(book, dataSheet, "DỮ LIỆU");
        const array = XLSX.write(book, { type: "array", bookType: "xlsx" });
        download(
          new Blob([array], {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          }),
          `MAU_NHAP_${safeFileName(schema.label).toUpperCase()}.xlsx`,
        );
      } catch (error) {
        toast("Không thể tạo mẫu Excel: " + error.message, "bad");
      }
    }
  }

  root.TPTUpgradeFeatures = Object.freeze({ create });
})(typeof globalThis !== "undefined" ? globalThis : window);
