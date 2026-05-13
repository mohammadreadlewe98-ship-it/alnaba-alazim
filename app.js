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

const regionFiles = {};

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
               accept=".xlsx,.xls,.csv"
               multiple
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

// ===== معالجة اختيار الملفات =====
function handleFileSelect(event, regionId) {
  const files = Array.from(event.target.files);
  files.forEach(file => processFile(file, regionId));
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
  const files = Array.from(event.dataTransfer.files);
  files.forEach(file => processFile(file, regionId));
}

// ===== قراءة الملف — يقرأ كل الأوراق =====
function processFile(file, regionId) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(e.target.result, { type: "binary" });

      // قراءة كل الأوراق وليس الأولى فقط
      let allRows = [];

      wb.SheetNames.forEach(sheetName => {
        const ws = wb.Sheets[sheetName];
        let rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

        // تجاهل الأوراق الفارغة
        if (rows.length === 0) return;

        // إضافة عمود المنطقة واسم الورقة والتاريخ لكل صف
        rows = rows.map(row => ({
          "المنطقة":    REGIONS[regionId].name,
          "اسم الورقة": sheetName,
          ...row,
          "تاريخ الرفع": new Date().toLocaleDateString("ar-SA")
        }));

        allRows = allRows.concat(rows);
      });

      if (allRows.length === 0) {
        alert(`⚠️ الملف فارغ: ${file.name}`);
        return;
      }

      // تحقق من التكرار
      if (regionFiles[regionId]) {
        const duplicate = regionFiles[regionId].find(f => f.name === file.name);
        if (duplicate) {
          if (!confirm(`الملف "${file.name}" مرفوع مسبقاً في هذه المنطقة.\nهل تريد استبداله؟`)) return;
          regionFiles[regionId] = regionFiles[regionId].filter(f => f.name !== file.name);
        }
      }

      if (!regionFiles[regionId]) regionFiles[regionId] = [];
      regionFiles[regionId].push({ name: file.name, rows: allRows });

      updateCardUI(regionId);
      updateGlobalStats();

    } catch (err) {
      alert(`❌ خطأ في قراءة الملف "${file.name}": ${err.message}`);
    }
  };
  reader.readAsBinaryString(file);
}

// ===== تحديث واجهة البطاقة =====
function updateCardUI(regionId) {
  const files = regionFiles[regionId] || [];
  const totalRows = files.reduce((sum, f) => sum + f.rows.length, 0);
  const fileCount = files.length;

  document.getElementById(`stats-${regionId}`).innerHTML =
    `<span class="files-badge">${fileCount} ملف</span>
     <span style="color:#7de8b0">${totalRows} سجل إجمالاً</span>`;

  const zone = document.getElementById(`zone-${regionId}`);
  zone.classList.add("success");
  zone.innerHTML = `
    <span class="drop-icon">➕</span>
    <p class="drop-text" style="color:#7de8b0">اضغط لإضافة ملف آخر</p>
    <input type="file" id="file-${regionId}"
           accept=".xlsx,.xls,.csv" multiple
           onchange="handleFileSelect(event, ${regionId})" />
  `;

  const list = document.getElementById(`files-list-${regionId}`);
  list.innerHTML = files.map((f, index) => `
    <div class="file-item">
      <span>📄 ${f.name}</span>
      <span>${f.rows.length} سجل</span>
      <button class="file-remove"
              onclick="removeFile(${regionId}, ${index})"
              title="حذف هذا الملف">✕</button>
    </div>
  `).join("");

  const regionBtn = document.getElementById(`btn-region-${regionId}`);
  regionBtn.style.display = fileCount > 0 ? "block" : "none";

  document.getElementById(`card-${regionId}`).classList.toggle("uploaded", fileCount > 0);
}

// ===== حذف ملف واحد من منطقة =====
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

// ===== إعادة بطاقة واحدة لوضعها الابتدائي =====
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

// ===== تحديث الإحصائيات العلوية =====
function updateGlobalStats() {
  const regionCount = Object.keys(regionFiles).length;
  const totalFiles  = Object.values(regionFiles).reduce((s, arr) => s + arr.length, 0);
  const totalRows   = Object.values(regionFiles).reduce((s, arr) =>
    s + arr.reduce((ss, f) => ss + f.rows.length, 0), 0);

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

  let merged = [];
  files.forEach(f => { merged = merged.concat(f.rows); });

  const ws = XLSX.utils.json_to_sheet(merged);
  ws["!cols"] = Object.keys(merged[0]).map(() => ({ wch: 22 }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, REGIONS[regionId].name);

  const today = new Date().toLocaleDateString("ar-SA").replace(/\//g, "-");
  XLSX.writeFile(wb, `${REGIONS[regionId].name}_${today}.xlsx`);
}

// ===== دمج كل المناطق وتصديرها =====
function mergeAllAndExport() {
  if (Object.keys(regionFiles).length === 0) {
    alert("⚠️ لم يتم رفع أي ملف");
    return;
  }

  let allRows = [];
  REGIONS.forEach(region => {
    if (regionFiles[region.id]) {
      regionFiles[region.id].forEach(f => {
        allRows = allRows.concat(f.rows);
      });
    }
  });

  // ورقة البيانات الكاملة
  const ws1 = XLSX.utils.json_to_sheet(allRows);
  ws1["!cols"] = Object.keys(allRows[0]).map(() => ({ wch: 22 }));

  // ورقة إحصائيات المناطق
  const summaryRows = REGIONS
    .filter(r => regionFiles[r.id])
    .map(r => ({
      "المنطقة":     r.name,
      "عدد الملفات": regionFiles[r.id].length,
      "عدد السجلات": regionFiles[r.id].reduce((s, f) => s + f.rows.length, 0)
    }));

  const ws2 = XLSX.utils.json_to_sheet(summaryRows);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws1, "البيانات الكاملة");
  XLSX.utils.book_append_sheet(wb, ws2, "إحصائيات المناطق");

  const today = new Date().toLocaleDateString("ar-SA").replace(/\//g, "-");
  XLSX.writeFile(wb, `النبأ_العظيم_${today}.xlsx`);

  document.getElementById("status-label").textContent = "تم التصدير ✅";
  renderPreview(allRows);
}

// ===== معاينة البيانات =====
let previewVisible = false;
function togglePreview() {
  const section = document.getElementById("preview-section");
  if (!previewVisible) {
    let allRows = [];
    Object.values(regionFiles).forEach(files =>
      files.forEach(f => { allRows = allRows.concat(f.rows); })
    );
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

  document.getElementById("preview-count").textContent = `${rows.length} سجل إجمالي`;
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
      يُعرض أول 50 سجل — الملف يحتوي على كامل البيانات (${rows.length} سجل)
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
