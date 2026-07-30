var TRACKING_SHEET_NAME = "قائمة المتابعة";

var PHCC_LIST = [
  { sector: "بارق", name: "Bariq PHCC" },
  { sector: "بارق", name: "Alkush" },
  { sector: "بارق", name: "Suhool PHCC" },
  { sector: "بارق", name: "Wadi Alkair PHCC" },
  { sector: "بارق", name: "Dhuha Sayala PHCC" },
  { sector: "بارق", name: "Thalooth Almandhar PHCC" },
  { sector: "بارق", name: "Sulaem PHCC" },
  { sector: "بارق", name: "Alqureha PHCC" },
  { sector: "بارق", name: "Jumat Rabia Almaqatra" },
  { sector: "بارق", name: "Almaslama PHCC" },
  { sector: "بارق", name: "South Bariq PHCC" },
  { sector: "بارق", name: "Khulsat Nusba PHCC" },
  { sector: "المجاردة", name: "Almajardah PHCC" },
  { sector: "المجاردة", name: "North Tharban" },
  { sector: "المجاردة", name: "Ahad Tharban PHCC" },
  { sector: "المجاردة", name: "Khat PHCC" },
  { sector: "المجاردة", name: "Altalalea PHCC" },
  { sector: "المجاردة", name: "Abss PHCC" },
  { sector: "المجاردة", name: "Khatba PHCC" },
  { sector: "المجاردة", name: "East Almajardah" },
  { sector: "المجاردة", name: "Al Ghaylan PHCC" }
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
  sheet.appendRow(["م", "القطاع", "اسم المركز الصحي (للمطابقة)", "عدد الردود المطابقة", "الحالة"]);
  sheet.getRange(1, 1, 1, 5).setFontWeight("bold");

  PHCC_LIST.forEach(function (center, i) {
    var row = i + 2;
    sheet.appendRow([i + 1, center.sector, center.name, "", ""]);
    sheet.getRange(row, 4).setFormula(
      "=COUNTIF('" + responsesSheetName + "'!B:B,\"*\"&C" + row + "&\"*\")"
    );
    sheet.getRange(row, 5).setFormula(
      "=IF(D" + row + ">0,\"✅ تم الاستلام\",\"⏳ بالانتظار\")"
    );
  });

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 5);
  return sheet;
}
