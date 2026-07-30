var TRACKING_SHEET_NAME = "قائمة المتابعة";

var PHCC_LIST = [
  { sector: "بارق", nameEn: "Bariq PHCC", nameAr: "مركز بارق الصحي" },
  { sector: "بارق", nameEn: "Alkush", nameAr: "مركز الكوش الصحي" },
  { sector: "بارق", nameEn: "Suhool PHCC", nameAr: "مركز سهول الصحي" },
  { sector: "بارق", nameEn: "Wadi Alkair PHCC", nameAr: "مركز وادي الخير الصحي" },
  { sector: "بارق", nameEn: "Dhuha Sayala PHCC", nameAr: "مركز ضحى سيالة الصحي" },
  { sector: "بارق", nameEn: "Thalooth Almandhar PHCC", nameAr: "مركز ثلوث المنذر الصحي" },
  { sector: "بارق", nameEn: "Sulaem PHCC", nameAr: "مركز سليم الصحي" },
  { sector: "بارق", nameEn: "Alqureha PHCC", nameAr: "مركز القريحاء الصحي" },
  { sector: "بارق", nameEn: "Jumat Rabia Almaqatra", nameAr: "مركز جمعة ربيعة المقطرة الصحي" },
  { sector: "بارق", nameEn: "Almaslama PHCC", nameAr: "مركز المسلمة الصحي" },
  { sector: "بارق", nameEn: "South Bariq PHCC", nameAr: "مركز جنوب بارق الصحي" },
  { sector: "بارق", nameEn: "Khulsat Nusba PHCC", nameAr: "مركز خلصة نصبة الصحي" },
  { sector: "المجاردة", nameEn: "Almajardah PHCC", nameAr: "مركز المجاردة الصحي" },
  { sector: "المجاردة", nameEn: "North Tharban", nameAr: "مركز شمال ثربان الصحي" },
  { sector: "المجاردة", nameEn: "Ahad Tharban PHCC", nameAr: "مركز أحد ثربان الصحي" },
  { sector: "المجاردة", nameEn: "Khat PHCC", nameAr: "مركز الخط الصحي" },
  { sector: "المجاردة", nameEn: "Altalalea PHCC", nameAr: "مركز التلاعة الصحي" },
  { sector: "المجاردة", nameEn: "Abss PHCC", nameAr: "مركز عبس الصحي" },
  { sector: "المجاردة", nameEn: "Khatba PHCC", nameAr: "مركز خطبة الصحي" },
  { sector: "المجاردة", nameEn: "East Almajardah", nameAr: "مركز شرق المجاردة الصحي" },
  { sector: "المجاردة", nameEn: "Al Ghaylan PHCC", nameAr: "مركز الغيلان الصحي" }
];

function doPost(e) {
  var responsesSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  ensureTrackingSheet(responsesSheet.getName());

  if (responsesSheet.getLastRow() === 0) {
    responsesSheet.appendRow([
      "التاريخ والوقت", "اسم المركز الصحي", "القطاع", "مدير المركز",
      "رقم التواصل", "البريد الإلكتروني", "منسق الاتفاقية", "رقم المنسق",
      "المركز المقترح", "مجال التعاون", "ملاحظات"
    ]);
  }

  var data = JSON.parse(e.postData.contents);

  responsesSheet.appendRow([
    new Date(),
    data["center-name"] || "",
    data.sector || "",
    data["manager-name"] || "",
    data.phone || "",
    data.email || "",
    data["coordinator-name"] || "",
    data["coordinator-phone"] || "",
    data["partner-center"] || "",
    (data.scope || []).join(" / "),
    data.notes || ""
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TRACKING_SHEET_NAME);
  if (!sheet) {
    var responsesSheet = ss.getSheets()[0];
    sheet = ensureTrackingSheet(responsesSheet.getName());
  }

  var lastRow = sheet.getLastRow();
  var stats = { total: 0, submitted: 0, sectors: {}, updatedAt: new Date().toISOString() };

  if (lastRow > 1) {
    var rows = sheet.getRange(2, 2, lastRow - 1, 5).getValues();
    rows.forEach(function (row) {
      var sector = row[0];
      var count = row[3];
      if (!stats.sectors[sector]) {
        stats.sectors[sector] = { total: 0, submitted: 0 };
      }
      stats.total++;
      stats.sectors[sector].total++;
      if (count > 0) {
        stats.submitted++;
        stats.sectors[sector].submitted++;
      }
    });
  }

  stats.percentage = stats.total ? Math.round((stats.submitted / stats.total) * 100) : 0;
  Object.keys(stats.sectors).forEach(function (name) {
    var s = stats.sectors[name];
    s.percentage = s.total ? Math.round((s.submitted / s.total) * 100) : 0;
  });

  return ContentService
    .createTextOutput(JSON.stringify(stats))
    .setMimeType(ContentService.MimeType.JSON);
}

// Run this once manually (Apps Script editor: select this function, then Run)
// to create the tracking tab immediately without waiting for a submission.
function setupTrackingSheet() {
  var responsesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  ensureTrackingSheet(responsesSheet.getName());
}

function ensureTrackingSheet(responsesSheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TRACKING_SHEET_NAME);
  if (sheet) return sheet;

  sheet = ss.insertSheet(TRACKING_SHEET_NAME);
  sheet.appendRow(["م", "القطاع", "الاسم بالإنجليزية", "اسم المركز الصحي (للمطابقة)", "عدد الردود المطابقة", "الحالة"]);
  sheet.getRange(1, 1, 1, 6).setFontWeight("bold");

  PHCC_LIST.forEach(function (center, i) {
    var row = i + 2;
    sheet.appendRow([i + 1, center.sector, center.nameEn, center.nameAr, "", ""]);
    sheet.getRange(row, 5).setFormula(
      "=COUNTIF('" + responsesSheetName + "'!B:B,\"*\"&D" + row + "&\"*\")"
    );
    sheet.getRange(row, 6).setFormula(
      "=IF(E" + row + ">0,\"✅ تم الاستلام\",\"⏳ بالانتظار\")"
    );
  });

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 6);
  return sheet;
}
