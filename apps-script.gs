function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "التاريخ والوقت", "اسم المركز الصحي", "القطاع", "مدير المركز",
      "رقم التواصل", "البريد الإلكتروني", "منسق الاتفاقية", "رقم المنسق",
      "المركز المقترح", "مجال التعاون", "ملاحظات"
    ]);
  }

  var data = JSON.parse(e.postData.contents);

  sheet.appendRow([
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
