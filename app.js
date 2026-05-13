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

const STUDENT_NAME_COLS = [
  "اسم الطالب", "الاسم الثلاثي", "اسم الطالبة", "الاسم"
];

const FIXED_COLS = [
  "المنطقة", "اسم الطالب", "اسم المدرس", "اسم الحلقة",
  "المركز", "المستوى", "_شهر_"
];

const HEADER_KEYWORDS = [
  "اسم الطالب", "الاسم الثلاثي", "اسم الطالبة",
  "اسم المدرس", "اسم المعلم", "المستوى", "التقدير", "م"
];

const regionFiles = {};

// ===== توحيد أسماء الأعمدة =====
function normalizeColumns(rows) {
  return rows.map(row => {
    const out = {};

    Object.entries(row).forEach(([key, val]) => {
      const k = key.trim();

      if (k === "م" || k === "رقم" || k === "#") {
        out["م"] = val;

      } else if (["اسم الطالب","الاسم الثلاثي","اسم الطالبة","الاسم"].includes(k)) {
        out["اسم الطالب"] = val;

      } else if (["اسم المدرس","اسم المعلم","المدرس","المعلم"].includes(k)) {
        out["اسم المدرس"] = val;

      } else if (
        k === "من" ||
        (k.includes("صفحات") && !k.includes("إلى") && !k.includes("خلال"))
      ) {
        if (!out["صفحة البداية"]) out["صفحة البداية"] = val;

      } else if (k === "إلى" || k === "عمود_7" || k === "عمود_9") {
        if (!out["صفحة النهاية"]) out["صفحة النهاية"] = val;

      } else if (k.includes("صفحات الحفظ خلال")) {
        if (!out["صفحة البداية"]) out["صفحة البداية"] = val;

      } else if (["التقدير","تقدير الحفظ","تقدير"].includes(k)) {
        out["التقدير"] = val;

      } else if (k.includes("أيام الحضور الفعلي")) {
        out["أيام الحضور"] = val;

      } else if (k.includes("أيام الغياب")) {
        out["أيام الغياب"] = val;

      } else if (k.includes("نسبة الحضور")) {
        out["نسبة الحضور"] = val;

      } else if (k.includes("أيام الداوام") || k.includes("أيام الدوام")) {
        out["أيام الدوام"] = val;

      } else if (k.includes("دروس المنهاج") || k.includes("عناوين دروس")) {
        out["دروس المنهاج"] = val;

      } else if (k.includes("الأجزاء المسبورة")) {
        out["الأجزاء المسبورة"] = val;

      } else if (k === "ملاحظات") {
        out["ملاحظات"] = val;

      } else if (k === "اسم الحلقة") {
        out["اسم الحلقة"] = val;

      } else if (k === "المركز") {
        out["المركز"] = val;

      } else if (k === "المستوى" || k.includes("المستوى (حسب")) {
        out["المستوى"] = val;

      } else if (!k.startsWith("عمود_") && k !== "_شهر_" && k !== "") {
        out[k] = val;
      }
    });

    if (row["_شهر_"]) out["_شهر_"] = row["_شهر_"];
    return out;
  });
}

// ===== إيجاد عمود اسم الطالب =====
function findStudentNameKey(row) {
  if (row["اسم الطالب"] !== undefined && String(row["اسم الطالب"]).trim() !== "")
    return "اسم الطالب";
  for (const col of STUDENT_NAME_COLS) {
    if (row.hasOwnProperty(col) && String(row[col]).trim() !== "") return col;
  }
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

// ===== استخراج اسم الشهر =====
function extractMonthLabel(fileName, sheetName) {
  const arabicMonths = [
    "يناير","فبراير","مارس","أبريل","مايو","يونيو",
    "يوليو","أغسطس","سبتمبر","أكتوبر","نوفمبر","ديسمبر",
    "محرم","صفر","ربيع","جمادى","رجب","شعبان","رمضان","شوال","ذو",
    "كانون","شباط","آذار","نيسان","أيار","حزيران",
    "تموز","آب","أيلول","تشرين","اذار","ايار"
  ];
  for (const src of [fileName, sheetName]) {
    for (const m of arabicMonths) {
      if (src.includes(m)) return m;
    }
  }
  return new Date().toLocaleDateString("ar-SA", { month: "long", year: "numeric" });
}

// ===== قراءة الملف مع البحث عن صف الرأس =====
function processFile(file, regionId) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(e.target.result, { type: "binary" });
      let allRows = [];
      const monthLabel = extractMonthLabel(file.name, wb.SheetNames[0]);

      wb.SheetNames.forEach(sheetName => {
        const ws = wb.Sheets[sheetName];

        const rawData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        if (rawData.length === 0) return;

        // البحث عن صف الرأس في أول 10 صفوف
        let headerRowIndex = 0;
        for (let i = 0; i < Math.min(rawData.length, 10); i++) {
          const rowStr = rawData[i].join(" ");
          const matchCount = HEADER_KEYWORDS.filter(k => rowStr.includes(k)).length;
          if (matchCount >= 2) { headerRowIndex = i; break; }
        }

        const headers = rawData[headerRowIndex].map((h, idx) => {
          const trimmed = String(h).trim();
          return trimmed !== "" ? trimmed : `عمود_${idx}`;
        });

        let rows = [];
        for (let i = headerRowIndex + 1; i < rawData.length; i++) {
          const rawRow = rawData[i];
          const rowObj = {};
          headers.forEach((h, idx) => {
            rowObj[h] = rawRow[idx] !== undefined ? rawRow[idx] : "";
          });

          const nonEmpty = Object.values(rowObj).filter(
            v => v !== "" && v !== null && v !== undefined
          );
          if (nonEmpty.length < 2) continue;

          const rowStr = Object.values(rowObj).join(" ");
          if (
            rowStr.includes("المجموع") ||
            rowStr.includes("الإجمالي") ||
            rowStr.includes("تقرير شهر") ||
            rowStr.includes("اسم الطالب") ||
            rowStr.includes("الاسم الثلاثي")
          ) continue;

          rowObj["_شهر_"] = monthLabel;
          rows.push(rowObj);
        }

        if (rows.length === 0) return;

        // توحيد الأعمدة
        rows = normalizeColumns(rows);
        allRows = allRows.concat(rows);
      });

      if (allRows.length === 0) {
        alert(`⚠️ لم يُعثر على بيانات في: ${file.name}`);
        return;
      }

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

// ===== الدمج الذكي: كل طالب سطر واحد =====
function mergeByStudentName(files, regionName) {
  if (!files || files.length === 0) return [];

  if (files.length === 1) {
    return files[0].rows.map(row => ({ "المنطقة": regionName, ...row }));
  }

  const byMonth = {};
  files.forEach(file => {
    if (!byMonth[file.month]) byMonth[file.month] = [];
    byMonth[file.month] = byMonth[file.month].concat(file.rows);
  });

  const months = Object.keys(byMonth);
  const studentMap = {};

  months.forEach(month => {
    byMonth[month].forEach(row => {
      const nameKey = findStudentNameKey(row);
      if (!nameKey) return;
      const studentName = String(row[nameKey]).trim();
      if (!studentName) return;

      if (!studentMap[studentName]) {
        studentMap[studentName] = { "المنطقة": regionName };
        FIXED_COLS.forEach(col => {
          if (row[col] !== undefined && row[col] !== "")
            studentMap[studentName][col] = row[col];
        });
      }

      Object.entries(row).forEach(([key, val]) => {
        if (FIXED_COLS.includes(key)) return;
        if (key === nameKey) return;
        studentMap[studentName][`${key} (${month})`] = val;
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
     <span style="color:#7de8b0">
       ${totalRows} سجل
       ${fileCount > 1 ? "— سيتم دمج الطلاب تلقائياً 🔀" : ""}
     </span>`;

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
      <button class="file-remove"
              onclick="removeFile(${regionId}, ${i})"
              title="حذف">✕</button>
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

// ===== تصدير منطقة واحدة =====
function exportRegionOnly(regionId) {
  const files = regionFiles[regionId];
  if (!files || files.length === 0) return;

  const merged = mergeByStudentName(files, REGIONS[regionId].name);
  if (merged.length === 0) { alert("⚠️ لا توجد بيانات"); return; }

  const ws = XLSX.utils.json_to_sheet(merged);
  ws["!cols"] = Object.keys(merged[0]).map(() => ({ wch: 22 }));

  const wbOut = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wbOut, ws, REGIONS[regionId].name);

  const today = new Date().toLocaleDateString("ar-SA").replace(/\//g, "-");
  XLSX.writeFile(wbOut, `${REGIONS[regionId].name}_${today}.xlsx`);
}

// ===== دمج كل المناطق — كل منطقة في ورقة مستقلة =====
function mergeAllAndExport() {
  if (Object.keys(regionFiles).length === 0) {
    alert("⚠️ لم يتم رفع أي ملف");
    return;
  }

  const wbOut = XLSX.utils.book_new();
  const summaryRows = [];
  let allRows = [];

  REGIONS.forEach(region => {
    if (!regionFiles[region.id]) return;

    const merged = mergeByStudentName(regionFiles[region.id], region.name);
    if (merged.length === 0) return;

    allRows = allRows.concat(merged);

    // ورقة مستقلة لكل منطقة
    const ws = XLSX.utils.json_to_sheet(merged);
    ws["!cols"] = Object.keys(merged[0]).map(() => ({ wch: 22 }));
    XLSX.utils.book_append_sheet(wbOut, ws, region.name.substring(0, 31));

    summaryRows.push({
      "المنطقة":     region.name,
      "عدد الملفات": regionFiles[region.id].length,
      "الأشهر":      regionFiles[region.id].map(f => f.month).join(" | "),
      "عدد الطلاب":  merged.length
    });
  });

  if (allRows.length === 0) { alert("⚠️ لا توجد بيانات"); return; }

  // ورقة كل المناطق مجتمعة
  const wsAll = XLSX.utils.json_to_sheet(allRows);
  wsAll["!cols"] = Object.keys(allRows[0]).map(() => ({ wch: 22 }));
  XLSX.utils.book_append_sheet(wbOut, wsAll, "كل المناطق");

  // ورقة الإحصائيات
  const wsSummary = XLSX.utils.json_to_sheet(summaryRows);
  XLSX.utils.book_append_sheet(wbOut, wsSummary, "الإحصائيات");

  const today = new Date().toLocaleDateString("ar-SA").replace(/\//g, "-");
  XLSX.writeFile(wbOut, `النبأ_العظيم_${today}.xlsx`);

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
    document.getElementById("preview-btn")
      .querySelector("span:last-child").textContent = "إخفاء المعاينة";
  } else {
    section.style.display = "none";
    document.getElementById("preview-btn")
      .querySelector("span:last-child").textContent = "معاينة البيانات";
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
    html += `<p style="text-align:center;color:var(--text-muted);
                        font-size:0.78rem;margin-top:12px">
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
