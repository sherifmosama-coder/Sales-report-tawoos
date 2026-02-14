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
    .setTitle('Sales Analytics Portal') // Generic Portal Title
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// SINGLE SOURCE OF TRUTH: Reading from Dashboard_Cache
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Dashboard_Cache");
  
  if (!sheet) return []; 

  const lastRow = sheet.getLastRow();
  // Adjust startRow to 2 if you have headers in Row 1
  if (lastRow < 2) return [];

  // Fetch Cols A to G (1 to 7)
  // Col A: Main Client, B: Branch, C: Type, D: Year, E: Qty, F: Orders, G: Max Date
  const rawData = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  
  const timeZone = ss.getSpreadsheetTimeZone() || "GMT";

  // Map rows to Objects (No Aggregation here - we send detailed data)
  const output = rawData.map(row => {
    const rawDate = row[6];
    let dateStr = "";
    
    // Date Formatting
    if (rawDate instanceof Date) {
      dateStr = Utilities.formatDate(rawDate, timeZone, "yyyy-MM-dd");
    } else if (typeof rawDate === 'string' && rawDate.length >= 10) {
      dateStr = rawDate.replace(/\//g, "-").slice(0, 10);
    }

    return {
      client: row[0],   // Main Client (Group Key)
      branch: row[1],   // Branch Name (Detail Key)
      type: row[2],
      year: row[3],
      qty: Number(row[4]) || 0,
      orders: Number(row[5]) || 0,
      lastTx: dateStr
    };
  }).filter(item => item.client && item.year); // Filter empty rows
  
  return output;
}

function parseNumber(val) {
  if (!val) return 0;
  if (val.toString().includes("#DIV/0!")) return 0;
  const clean = val.toString().replace(/,/g, '');
  const num = parseFloat(clean);
  return isNaN(num) ? 0 : num;
}
