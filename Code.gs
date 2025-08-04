function doGet() {
  const ss = SpreadsheetApp.openById("1E2KH0fWsOd7Gywbx8qPzNyilNAE1EJWwWGffBAzf0VM");
  const sheet = ss.getSheetByName("data");

  let html = "<html><head><title>data 시트</title></head><body>";

  if (sheet) {
    const data = sheet.getDataRange().getValues();
    html += "<h2>data 시트</h2>";
    html += "<table border='1' cellpadding='5' cellspacing='0'>";
    data.forEach((row, index) => {
      html += "<tr>";
      row.forEach(cell => {
        html += index === 0
          ? `<th>${cell}</th>`
          : `<td>${cell}</td>`;
      });
      html += "</tr>";
    });
    html += "</table>";
  } else {
    html += "<p><strong>data</strong> 시트를 찾을 수 없습니다.</p>";
  }

  html += "</body></html>";
  return ContentService.createHtmlOutput(html);
}
