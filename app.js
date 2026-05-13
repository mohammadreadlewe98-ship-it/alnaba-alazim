// ===== أسماء المناطق =====
const REGIONS = [
  { id: 0, name: "المنطقة الأولى"  },
  { id: 1, name: "المنطقة الثانية" },
  { id: 2, name: "المنطقة الثالثة" },
  { id: 3, name: "المنطقة الرابعة" },
  { id: 4, name: "المنطقة الخامسة" },
  { id: 5, name: "المنطقة السادسة" },
  { id: 6, name: "المنطقة السابعة" }
];

// عمود اسم الطالب — عدّله إذا كان اسم العمود مختلفاً في ملفاتك
const STUDENT_NAME_COLS = [
  "اسم الطالب",
  "الاسم الثلاثي",
  "اسم الطالبة",
  "الاسم"
];

const regionFiles = {};

// ===== إيجاد عمود اسم الطالب في أي صف =====
function findStudentNameKey(row) {
  for (const col of STUDENT_NAME_COLS) {
    if (row.hasOwnProperty(col) && row[col] !== "") return col;
  }
  // إذا لم يُوجد، أعد أول مفتاح غير فارغ
  return Object.keys(row).find(k => row[k] !== "") || null;
}

// ===== بناء البطاقات =====
document.addEventListener("DOMContentLoaded", () => {
  document.getElementById("year").textContent = new Date().getFullYear();
  const grid = document.getElementById("regions-grid");

  REGIONS.forEach(region => {
    const card = document.createElement("div");
    card.className = "region-card";
    card.id = `card-${region.id}`;
    card.innerHTML = `
      <div class="card-header">
        <span class="card-title">${region.name}</span>
        <span class="card-number">${region.id + 1}</span>
      </div>
      <p class="card-stats" id="stats-${region.id}">لم يتم رفع ملف بعد</p>
      <div class="drop-zone" id="zone-${region.id}"
           onclick="document.getElementById('file-${region.id}').click()"
           ondragover="handleDragOver(event, ${region.id})"
           ondragleave="handleDragLeave(event, ${region.id})"
           ondrop="handleDrop(event, ${region.id})">
        <span class="drop-icon">📁</span>
        <p class="drop-text">
          اضغط لإضافة ملف أو أكثر<br>
          <span style="color:rgba(255,255,255,0.3)">.xlsx / .xls / .csv</span>
        </p>
        <input type="file" id="file-${region.id}"
               accept=".xlsx,.xls,.csv" multiple
               onchange="handleFileSelect(event, ${region.id})" />
      </div>
      <div class="files-list" id="files-list-${region.id}"></div>
      <button class="btn-region-merge"
              id="btn-region-${region.id}"
              onclick="exportRegionOnly(${region.id})">
        📥 تصدير هذه المنطقة منفردةً
      </button>
    `;
    grid.appendChild(card);
  });
});

// ===== السحب والإفلات =====
function handleFileSelect(event, regionId) {
  Array.from(event.target.files).forEach(f => processFile(f, regionId));
  event.target.value = "";
}
function handleDragOver(event, regionId) {
  event.preventDefault();
  document.getElementById(`zone-${regionId}`).classList.add("drag-over");
}
function handleDragLeave(event, regionId) {
  document.getElementById(`zone-${regionId}`).classList.remove("drag-over");
}
function handleDrop(event, regionId) {
  event.preventDefault();
  document.getElementById(`zone-${regionId}`).classList.remove("drag-over");
  Array.from(event.dataTransfer.files).forEach(f => processFile(f, regionId));
}

// ===== قراءة الملف مع تنظيف الأعمدة =====
function processFile(file, regionId) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(e.target.result, { type: "binary" });
      let allRows = [];

      // استخراج اسم الشهر من اسم الملف أو الورقة
      const monthLabel = extractMonthLabel(file.name, wb.SheetNames[0]);

      wb.SheetNames.forEach(sheetName => {
        const ws = wb.Sheets[sheetName];
        let rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
        if (rows.length === 0) return;

        // تنظيف الأعمدة الفارغة و __EMPTY
        rows = rows.map(row => {
          const clean = {};
          Object.entries(row).forEach(([k, v]) => {
            if (k.startsWith("__EMPTY") || k.trim() === "") return;
            clean[k.trim()] = v;
          });
          return clean;
        }).filter(row =>
          Object.values(row).filter(v => v !== "" && v !== null).length > 2
        );

        if (rows.length === 0) return;

        // إضافة عمود الشهر لكل صف
        rows = rows.map(row => ({ ...row, "_شهر_": monthLabel }));
        allRows = allRows.concat(rows);
      });

      if (allRows.length === 0) {
        alert(`⚠️ لم يُعثر على بيانات في: ${file.name}`);
        return;
      }

      // التحقق من التكرار
      if (regionFiles[regionId]) {
        const dup = regionFiles[regionId].find(f => f.name === file.name);
        if (dup) {
          if (!confirm(`الملف "${file.name}" مرفوع مسبقاً.\nهل تريد استبداله؟`)) return;
          regionFiles[regionId] = regionFiles[regionId].filter(f => f.name !== file.name);
        }
      }

      if (!regionFiles[regionId]) regionFiles[regionId] = [];
      regionFiles[regionId].push({
        name: file.name,
        rows: allRows,
        month: monthLabel,
        regionName: REGIONS[regionId].name
      });

      updateCardUI(regionId);
      updateGlobalStats();

    } catch (err) {
      alert(`❌ خطأ في قراءة "${file.name}": ${err.message}`);
    }
  };
  reader.readAsBinaryString(file);
}

// ===== استخراج اسم الشهر من اسم الملف أو الورقة =====
function extractMonthLabel(fileName, sheetName) {
  const arabicMonths = [
    "يناير","فبراير","مارس","أبريل","مايو","يونيو",
    "يوليو","أغسطس","سبتمبر","أكتوبر","نوفمبر","ديسمبر",
    "محرم","صفر","ربيع","جمادى","رجب","شعبان","رمضان","شوال","ذو",
    "كانون","شباط","آذار","نيسان","أيار","حزيران",
    "تموز","آب","أيلول","تشرين","اذار","ايار"
  ];

  const sources = [fileName, sheetName];
  for (const src of sources) {
    for (const m of arabicMonths) {
      if (src.includes(m)) return m;
    }
  }

  // إذا لم يُعثر على شهر، استخدم تاريخ اليوم
  return new Date().toLocaleDateString("ar-SA", { month: "long", year: "numeric" });
}

// ===== الدمج الذكي: مقارنة الطلاب وتوسيع الأعمدة أفقياً =====
function mergeByStudentName(files, regionName) {
  if (!files || files.length === 0) return [];

  // إذا كان ملف واحد فقط، أعده كما هو مع إضافة المنطقة
  if (files.length === 1) {
    return files[0].rows.map(row => ({
      "المنطقة": regionName,
      ...row
    }));
  }

  // ===== دمج متعدد الملفات بمقارنة اسم الطالب =====

  // خطوة 1: جمع كل الملفات وتصنيفها حسب الشهر
  const byMonth = {};
  files.forEach(file => {
    const month = file.month;
    if (!byMonth[month]) byMonth[month] = [];
    byMonth[month] = byMonth[month].concat(file.rows);
  });

  const months = Object.keys(byMonth);

  // خطوة 2: بناء قاموس الطلاب
  // المفتاح: اسم الطالب | القيمة: بياناته من كل شهر
  const studentMap = {};

  // الأعمدة الثابتة (لا تتكرر مع كل شهر)
  const FIXED_COLS = [
    "المنطقة", "اسم الطالب", "الاسم الثلاثي", "اسم الطالبة",
    "الاسم", "اسم المدرس", "اسم المعلم", "اسم الحلقة",
    "المركز", "المستوى", "_شهر_"
  ];

  months.forEach(month => {
    byMonth[month].forEach(row => {
      // إيجاد اسم الطالب
      const nameKey = findStudentNameKey(row);
      if (!nameKey) return;
      const studentName = String(row[nameKey]).trim();
      if (!studentName) return;

      if (!studentMap[studentName]) {
        // أول ظهور لهذا الطالب: نحفظ بياناته الثابتة
        studentMap[studentName] = {
          "المنطقة": regionName,
          [nameKey]:  studentName,
        };

        // حفظ الأعمدة الثابتة
        FIXED_COLS.forEach(col => {
          if (row[col] !== undefined && row[col] !== "") {
            studentMap[studentName][col] = row[col];
          }
        });
      }

      // إضافة بيانات هذا الشهر كأعمدة جديدة بلاحقة الشهر
      Object.entries(row).forEach(([key, val]) => {
        if (FIXED_COLS.includes(key)) return; // تجاهل الأعمدة الثابتة
        if (key === nameKey) return;
        const newKey = `${key} (${month})`;
        studentMap[studentName][newKey] = val;
      });
    });
  });

  return Object.values(studentMap);
}

// ===== تحديث واجهة البطاقة =====
function updateCardUI(regionId) {
  const files = regionFiles[regionId] || [];
  const totalRows = files.reduce((sum, f) => sum + f.rows.length, 0);
  const fileCount = files.length;

  document.getElementById(`stats-${regionId}`).innerHTML =
    `<span class="files-badge">${fileCount} ملف</span>
     <span style="color:#7de8b0">${totalRows} سجل — ${fileCount > 1 ? "سيتم دمج الطلاب تلقائياً 🔀" : ""}  </span>`;

  const zone = document.getElementById(`zone-${regionId}`);
  zone.classList.add("success");
  zone.innerHTML = `
    <span class="drop-icon">➕</span>
    <p class="drop-text" style="color:#7de8b0">اضغط لإضافة شهر آخر</p>
    <input type="file" id="file-${regionId}"
           accept=".xlsx,.xls,.csv" multiple
           onchange="handleFileSelect(event, ${regionId})" />
  `;

  const list = document.getElementById(`files-list-${regionId}`);
  list.innerHTML = files.map((f, i) => `
    <div class="file-item">
      <span>📄 ${f.name}</span>
      <span style="color:var(--gold-light)">${f.month}</span>
      <span>${f.rows.length} طالب</span>
      <button class="file-remove" onclick="removeFile(${regionId}, ${i})" title="حذف">✕</button>
    </div>
  `).join("");

  document.getElementById(`btn-region-${regionId}`).style.display =
    fileCount > 0 ? "block" : "none";
  document.getElementById(`card-${regionId}`).classList.toggle("uploaded", fileCount > 0);
}

// ===== حذف ملف =====
function removeFile(regionId, fileIndex) {
  regionFiles[regionId].splice(fileIndex, 1);
  if (regionFiles[regionId].length === 0) {
    delete regionFiles[regionId];
    resetCard(regionId);
  } else {
    updateCardUI(regionId);
  }
  updateGlobalStats();
}

// ===== إعادة تعيين بطاقة =====
function resetCard(regionId) {
  document.getElementById(`stats-${regionId}`).textContent = "لم يتم رفع ملف بعد";
  document.getElementById(`files-list-${regionId}`).innerHTML = "";
  document.getElementById(`btn-region-${regionId}`).style.display = "none";
  document.getElementById(`card-${regionId}`).classList.remove("uploaded");
  const zone = document.getElementById(`zone-${regionId}`);
  zone.classList.remove("success");
  zone.innerHTML = `
    <span class="drop-icon">📁</span>
    <p class="drop-text">اضغط لإضافة ملف أو أكثر<br>
    <span style="color:rgba(255,255,255,0.3)">.xlsx / .xls / .csv</span></p>
    <input type="file" id="file-${regionId}"
           accept=".xlsx,.xls,.csv" multiple
           onchange="handleFileSelect(event, ${regionId})" />
  `;
}

// ===== تحديث الإحصائيات =====
function updateGlobalStats() {
  const regionCount = Object.keys(regionFiles).length;
  const totalFiles  = Object.values(regionFiles).reduce((s, a) => s + a.length, 0);
  const totalRows   = Object.values(regionFiles).reduce((s, a) =>
    s + a.reduce((ss, f) => ss + f.rows.length, 0), 0);

  document.getElementById("uploaded-count").textContent = regionCount;
  document.getElementById("total-files").textContent    = totalFiles;
  document.getElementById("total-records").textContent  = totalRows;
  document.getElementById("status-label").textContent   =
    regionCount === 7 ? "مكتمل ✅" : regionCount > 0 ? "جزئي ⏳" : "جاهز";

  document.getElementById("merge-btn").disabled = regionCount === 0;
  document.getElementById("preview-btn").style.display =
    regionCount > 0 ? "inline-flex" : "none";
}

// ===== تصدير منطقة واحدة فقط =====
function exportRegionOnly(regionId) {
  const files = regionFiles[regionId];
  if (!files || files.length === 0) return;

  const merged = mergeByStudentName(files, REGIONS[regionId].name);
  if (merged.length === 0) { alert("⚠️ لا توجد بيانات"); return; }

  const ws = XLSX.utils.json_to_sheet(merged);
  ws["!cols"] = Object.keys(merged[0]).map(() => ({ wch: 22 }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, REGIONS[regionId].name);

  const today = new Date().toLocaleDateString("ar-SA").replace(/\//g, "-");
  XLSX.writeFile(wb, `${REGIONS[regionId].name}_${today}.xlsx`);
}

// ===== دمج كل المناطق =====
function mergeAllAndExport() {
  if (Object.keys(regionFiles).length === 0) {
    alert("⚠️ لم يتم رفع أي ملف");
    return;
  }

  let allRows = [];
  const summaryRows = [];

  REGIONS.forEach(region => {
    if (!regionFiles[region.id]) return;
    const merged = mergeByStudentName(regionFiles[region.id], region.name);
    allRows = allRows.concat(merged);

    summaryRows.push({
      "المنطقة":      region.name,
      "عدد الملفات":  regionFiles[region.id].length,
      "الأشهر":       regionFiles[region.id].map(f => f.month).join(" | "),
      "عدد الطلاب":   merged.length
    });
  });

  if (allRows.length === 0) { alert("⚠️ لا توجد بيانات"); return; }

  const ws1 = XLSX.utils.json_to_sheet(allRows);
  ws1["!cols"] = Object.keys(allRows[0]).map(() => ({ wch: 22 }));

  const ws2 = XLSX.utils.json_to_sheet(summaryRows);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws1, "البيانات الكاملة");
  XLSX.utils.book_append_sheet(wb, ws2, "إحصائيات المناطق");

  const today = new Date().toLocaleDateString("ar-SA").replace(/\//g, "-");
  XLSX.writeFile(wb, `النبأ_العظيم_${today}.xlsx`);

  document.getElementById("status-label").textContent = "تم التصدير ✅";
  renderPreview(allRows);
}

// ===== المعاينة =====
let previewVisible = false;
function togglePreview() {
  const section = document.getElementById("preview-section");
  if (!previewVisible) {
    let allRows = [];
    Object.entries(regionFiles).forEach(([id, files]) => {
      const merged = mergeByStudentName(files, REGIONS[parseInt(id)].name);
      allRows = allRows.concat(merged);
    });
    renderPreview(allRows);
    section.style.display = "block";
    section.scrollIntoView({ behavior: "smooth" });
    document.getElementById("preview-btn").querySelector("span:last-child").textContent = "إخفاء المعاينة";
  } else {
    section.style.display = "none";
    document.getElementById("preview-btn").querySelector("span:last-child").textContent = "معاينة البيانات";
  }
  previewVisible = !previewVisible;
}

function renderPreview(rows) {
  if (!rows.length) return;
  const preview = rows.slice(0, 50);
  const headers = Object.keys(preview[0]);

  document.getElementById("preview-count").textContent = `${rows.length} طالب إجمالي`;
  document.getElementById("preview-section").style.display = "block";

  let html = `<table><thead><tr>
    ${headers.map(h => `<th>${h}</th>`).join("")}
  </tr></thead><tbody>`;
  preview.forEach(row => {
    html += `<tr>${headers.map(h => `<td>${row[h] ?? ""}</td>`).join("")}</tr>`;
  });
  html += `</tbody></table>`;
  if (rows.length > 50) {
    html += `<p style="text-align:center;color:var(--text-muted);font-size:0.78rem;margin-top:12px">
      يُعرض أول 50 طالب — الملف يحتوي على كامل البيانات (${rows.length} طالب)
    </p>`;
  }
  document.getElementById("preview-table-container").innerHTML = html;
}

// ===== إعادة تعيين كل شيء =====
function resetAll() {
  if (!confirm("⚠️ هل تريد حذف كل البيانات والبدء من جديد؟")) return;
  Object.keys(regionFiles).forEach(k => delete regionFiles[k]);
  REGIONS.forEach(r => resetCard(r.id));
  document.getElementById("preview-section").style.display = "none";
  document.getElementById("preview-btn").style.display = "none";
  previewVisible = false;
  updateGlobalStats();
}
