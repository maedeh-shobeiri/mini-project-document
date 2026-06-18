// دریافت عناصر
const fileInput = document.getElementById("fileInput");
const btnAdd = document.getElementById("btn-add");
const tableBody = document.getElementById("tableBody");

// باز کردن فایل‌چوژر با کلیک روی دکمه
btnAdd.addEventListener("click", () => fileInput.click());

// خواندن فایل پس از انتخاب
fileInput.addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: "array" });

    // فرض بر این است که داده‌ها در اولین شیت هستند
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(firstSheet);

    renderTable(jsonData);
  };
  reader.readAsArrayBuffer(file);
});

// تابع نمایش داده‌ها در جدول
function renderTable(data) {
  const tableBody = document.querySelector(".data-table tbody");
  tableBody.innerHTML = ""; // پاک کردن ردیف‌های قبلی

  data.forEach((row) => {
    const tr = document.createElement("tr");

    // نام ستون‌ها باید دقیقاً مطابق اکسل باشد
    const radif = row["ردیف"] || "";
    const hesab = row["حساب"] || "";
    const sharh = row["شرح"] || "";
    const type = row["بدهکار / بستاندار"] || ""; // ستون مربوطه با همان غلط املایی فایل اکسل

    tr.innerHTML = `
            <td>${radif}</td>
            <td>${hesab}</td>
            <td>${sharh}</td>
            <td>${type}</td>
        `;
    tableBody.appendChild(tr);
  });
}

async function save() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF("p", "mm", "a4");

  // لود فونت (حتماً مسیر فایل در پروژه درست باشد)
  const fontResponse = await fetch("font/IRANYekanXFaNum-Regular.ttf");
  const fontBuffer = await fontResponse.arrayBuffer();
  let binary = "";
  const bytes = new Uint8Array(fontBuffer);
  for (let i = 0; i < bytes.byteLength; i++)
    binary += String.fromCharCode(bytes[i]);

  doc.addFileToVFS("IRANYekan.ttf", btoa(binary));
  doc.addFont("IRANYekan.ttf", "IRANYekan", "normal");
  doc.setFont("IRANYekan");

  // گرفتن مقادیر فیلدها
  const unit = document.getElementById("unitRegister").value;
  const type = document.getElementById("documentType").value;
  const refNo = document.getElementById("refNumber").value;
  const desc = document.getElementById("docDescription").value;

  // نوشتن هدر سند (RTL)
  doc.setFontSize(14);
  doc.text(`واحد ثبت: ${unit}`, 190, 20, { align: "right" });
  doc.text(`نوع سند: ${type}`, 190, 30, { align: "right" });
  doc.text(`شماره عطف: ${refNo}`, 190, 40, { align: "right" });
  doc.text(`شرح: ${desc}`, 190, 50, { align: "right" });

  // استخراج داده‌ها از جدول HTML به صورت آرایه (برای جلوگیری از خرابی فونت در تبدیل مستقیم)
  const tableRows = [];
  const rows = document.querySelectorAll(".data-table tbody tr");
  rows.forEach((row) => {
    const cells = row.querySelectorAll("td");
    if (cells.length > 1) {
      tableRows.push([
        cells[3].innerText, // بدهکار/بستانکار
        cells[2].innerText, // شرح
        cells[1].innerText, // حساب
        cells[0].innerText, // ردیف
      ]);
    }
  });

  // رسم جدول
  doc.autoTable({
    startY: 65,
    head: [["بدهکار / بستانکار", "شرح", "حساب", "ردیف"]],
    body: tableRows,
    styles: { font: "IRANYekan", halign: "center" },
    headStyles: { fillColor: [93, 149, 210], halign: "center" },
    direction: "rtl", // مهم برای چینش فارسی
  });

  doc.save("Accounting_Document.pdf");
}

if (typeof XLSX === "undefined") {
  console.error("کتابخانه XLSX پیدا نشد! فایل JS را چک کنید.");
} else {
  console.log("کتابخانه XLSX با موفقیت لود شد.");
}
