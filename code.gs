function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Manga Dashboard')
      .addItem('View Dashboard', 'showDashboard')
      .addToUi();
}

// Function to open the dashboard as a popup inside the sheet
function showDashboard() {
  var html = HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setWidth(1200)
      .setHeight(850);
  SpreadsheetApp.getUi().showModalDialog(html, 'Manga Sales Analytics');
}

// This function is required if you want to deploy as a Web App URL
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Manga Sales Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Pulls data from the sheet and sends it to the HTML page
function getMangaData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1);
  
  return rows.map(function(row) {
    var obj = {};
    headers.forEach(function(header, i) {
      // Create clean keys: "Manga series" becomes "manga_series"
      var key = header.toString().toLowerCase().replace(/[^a-z0-9]/g, '_');
      obj[key] = row[i];
    });
    return obj;
  }).filter(item => item.manga_series); // Ignore empty rows
}