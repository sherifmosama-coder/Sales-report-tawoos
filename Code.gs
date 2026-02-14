function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Sales Dashboard')
    .addItem('Launch Dashboard', 'openDashboard')
    .addToUi();
}

function openDashboard() {
  // 🔴 REPLACE THIS with your actual Web App URL
  var webAppUrl = "https://script.google.com/macros/s/AKfycbzIjjcSJZeAdeW9HDz__thoM2V2ACW-LQMGlaEa_RdtzyEgkXV9BhMMopA_RNF794o8/exec";
  
  // We create a small HTML button that opens the link in a new tab (_blank)
  // The 'onclick="google.script.host.close()"' part closes the small popup immediately after you click.
  var html = HtmlService.createHtmlOutput(
    '<html>' +
    '<head><style>' +
    'body { display:flex; flex-direction:column; justify-content:center; align-items:center; height:100vh; margin:0; font-family: sans-serif; background-color:#f8f9fa; }' +
    '.btn { background-color: #2563eb; color: white; padding: 12px 24px; text-decoration: none; border-radius: 8px; font-weight: bold; font-size: 16px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); transition: background 0.2s; }' +
    '.btn:hover { background-color: #1d4ed8; }' +
    '</style></head>' +
    '<body>' +
    '<a href="' + webAppUrl + '" target="_blank" class="btn" onclick="google.script.host.close()">' +
    '🚀 Open Dashboard in New Tab' +
    '</a>' +
    '<p style="color:#64748b; margin-top:10px; font-size:12px;">Click above to launch the full screen view</p>' +
    '</body>' +
    '</html>'
  )
  .setWidth(400)
  .setHeight(200);

  SpreadsheetApp.getUi().showModalDialog(html, 'Launch Dashboard');
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Advanced Sales Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDataFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("ALL YEARS");
  const values = sheet.getDataRange().getDisplayValues(); 
  const dataRows = values.slice(3); 
  
  const cleanData = [];
  const years = [2021, 2022, 2023, 2024, 2025, 2026];
  
  dataRows.forEach(row => {
    const clientName = row[1]; 
    const clientType = row[2]; 
    const lastDate   = row[3]; 
    
    if (!clientName) return;

    // Create a record for each year
    years.forEach((year, index) => {
      const colIdx = 4 + (index * 3);
      const totalQty = parseNumber(row[colIdx]);
      const orderCount = parseNumber(row[colIdx+1]);
      const avgQty = parseNumber(row[colIdx+2]);
      
      // We push ALL years (even with 0 sales) to help with "Declining" calculation, 
      // but we will hide empty rows in the UI later if needed.
      if (clientName) {
        cleanData.push({
          client: clientName,
          type: clientType || "Other",
          lastTx: lastDate,
          year: year,
          qty: totalQty,
          orders: orderCount,
          avg: avgQty
        });
      }
    });
  });
  
  return cleanData;
}

function parseNumber(val) {
  if (!val) return 0;
  if (val.toString().includes("#DIV/0!")) return 0;
  const clean = val.toString().replace(/,/g, '');
  const num = parseFloat(clean);
  return isNaN(num) ? 0 : num;
}
