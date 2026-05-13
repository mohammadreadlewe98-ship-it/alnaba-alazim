// ===== أسماء المناطق — عدّلها حسب محافظاتك =====
const REGIONS = [
  { id: 0, name: "المنطقة الأولى"  },
  { id: 1, name: "المنطقة الثانية" },
  { id: 2, name: "المنطقة الثالثة" },
  { id: 3, name: "المنطقة الرابعة" },
  { id: 4, name: "المنطقة الخامسة" },
  { id: 5, name: "المنطقة السادسة" },
  { id: 6, name: "المنطقة السابعة" }
];

// تخزين بيانات كل منطقة بعد الرفع
const regionData = {};

// ===== بناء البطاقات عند تحميل الصفحة =====
document.addEventListener("DOMContentLoaded", () => {
  // سنة الفوتر
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
        <p class="drop-text">اضغط لاختيار ملف<br><span style="color:rgba(255,255,255,0.3)">.xlsx / .xls / .csv</span></p>
        <input type="file" id="file-${region.id}"
               accept=".xlsx,.xls,.csv"
               onchange="handleFileSelect(event, ${region.id})" />
      </div>
    `;
    grid.appendChild(card);
  });
});

// ===== معالجة اختيار الملف =====
function handleFileSelect(event, regionId) {
  const file = event.target.files[0];
  if (file) processFile(file, regionId);
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
  const file = event.dataTransfer.files[0];
  if (file) processFile(file, regionId);
}

// ===== قراءة ومعالجة الملف =====
function processFile(file, regionId) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(e.target.result, { type: "binary" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      let rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

      if (rows.length === 0) {
        alert(`⚠️ الملف فارغ: ${file.name}`);
        return;
      }

      // إضافة المنطقة والتاريخ لكل صف
      rows = rows.map(row => ({
        "المنطقة": REGIONS[regionId].name,
        ...row,
        "تاريخ الرفع": new Date().toLocaleDateString("ar-SA")
      }));

      regionData[regionId] = rows;
      updateCardUI(regionId, file.name, rows.length);
      updateStats();

    } catch (err) {
      alert(`❌ خطأ في قراءة الملف: ${err.message}`);
    }
  };
  reader.readAsBinaryString(file);
}

// ===== تحديث واجهة البطاقة =====
function updateCardUI(regionId, fileName, rowCount) {
  document.getElementById(`stats-${regionId}`).innerHTML =
    `<span style="color:#7de8b0">✅ ${rowCount} سجل</span> — <span style="opacity:0.7">${fileName}</span>`;

  const zone = document.getElementById(`zone-${regionId}`);
  zone.classList.remove("drag-over");
  zone.classList.add("success");
  zone.innerHTML = `
    <span class="drop-icon">✅</span>
    <p class="drop-text" style="color:#7de8b0; font-weight:700">${fileName}</p>
    <p class="drop-text">${rowCount} سجل محمّل</p>
    <p class="drop-text" style="font-size:0.7rem; margin-top:4px; opacity:0.5">اضغط لتغيير الملف</p>
    <input type="file" id="file-${regionId}"
           accept=".xlsx,.xls,.csv"
           onchange="handleFileSelect(event, ${regionId})" />
  `;
  document.getElementById(`card-${regionId}`).classList.add("uploaded");
}

// ===== تحديث الإحصائيات العلوية =====
function updateStats() {
  const count = Object.keys(regionData).length;
  const total = Object.values(regionData).reduce((sum, rows) => sum + rows.length, 0);

  document.getElementById("uploaded-count").textContent = count;
  document.getElementById("total-records").textContent = total;
  document.getElementById("status-label").textContent =
    count === 7 ? "مكتمل ✅" : count > 0 ? "جزئي ⏳" : "جاهز";

  // تفعيل زر الدمج
  document.getElementById("merge-btn").disabled = count === 0;

  // إظهار زر المعاينة
  if (count > 0) {
    document.getElementById("preview-btn").style.display = "inline-flex";
  }
}

// ===== الدمج والتصدير =====
function mergeAndExport() {
  if (Object.keys(regionData).length === 0) {
    alert("⚠️ لم يتم رفع أي ملف");
    return;
  }

  // دمج كل البيانات
  let allRows = [];
  Object.values(regionData).forEach(rows => { allRows = allRows.concat(rows); });

  // إنشاء ورقة البيانات الكاملة
  const ws1 = XLSX.utils.json_to_sheet(allRows);
  ws1["!cols"] = Object.keys(allRows[0]).map(() => ({ wch: 22 }));

  // إنشاء ورقة الإحصائيات
  const summaryRows = REGIONS
    .filter(r => regionData[r.id])
    .map(r => ({
      "المنطقة":        r.name,
      "عدد السجلات":    regionData[r.id].length,
      "تاريخ الرفع":    regionData[r.id][0]["تاريخ الرفع"] || ""
    }));

  const ws2 = XLSX.utils.json_to_sheet(summaryRows);

  // تجميع الكتاب وتصديره
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws1, "البيانات الكاملة");
  XLSX.utils.book_append_sheet(wb, ws2, "إحصائيات المناطق");

  const today = new Date().toLocaleDateString("ar-SA").replace(/\//g, "-");
  XLSX.writeFile(wb, `النبأ_العظيم_${today}.xlsx`);

  document.getElementById("status-label").textContent = "تم التصدير ✅";

  // تحديث المعاينة تلقائياً
  renderPreview(allRows);
}

// ===== تبديل ظهور المعاينة =====
let previewVisible = false;
function togglePreview() {
  const section = document.getElementById("preview-section");
  if (!previewVisible) {
    let allRows = [];
    Object.values(regionData).forEach(rows => { allRows = allRows.concat(rows); });
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

// ===== رسم جدول المعاينة =====
function renderPreview(rows) {
  if (!rows.length) return;
  const preview = rows.slice(0, 50);
  const headers = Object.keys(preview[0]);

  document.getElementById("preview-count").textContent =
    `${rows.length} سجل إجمالي`;

  let html = `<table><thead><tr>
    ${headers.map(h => `<th>${h}</th>`).join("")}
  </tr></thead><tbody>`;

  preview.forEach(row => {
    html += `<tr>${headers.map(h => `<td>${row[h] ?? ""}</td>`).join("")}</tr>`;
  });

  html += `</tbody></table>`;
  if (rows.length > 50) {
    html += `<p style="text-align:center;color:var(--text-muted);font-size:0.78rem;margin-top:12px">
      يُعرض أول 50 سجل — الملف المُصدَّر يحتوي على كامل البيانات (${rows.length} سجل)
    </p>`;
  }

  document.getElementById("preview-section").style.display = "block";
  document.getElementById("preview-table-container").innerHTML = html;
}

// ===== إعادة تعيين =====
function resetAll() {
  if (!confirm("⚠️ هل تريد حذف كل البيانات والبدء من جديد؟")) return;
  Object.keys(regionData).forEach(k => delete regionData[k]);
  document.getElementById("regions-grid").innerHTML = "";
  document.getElementById("preview-section").style.display = "none";
  document.getElementById("preview-btn").style.display = "none";
  previewVisible = false;
  document.dispatchEvent(new Event("DOMContentLoaded"));
  updateStats();
}