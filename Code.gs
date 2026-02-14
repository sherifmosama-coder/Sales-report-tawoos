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
    .setTitle('Sales Dashboard')
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
    // UPDATED: Shifted indices to ignore Column B. Data now starts at Column C (index 2).
    const clientName = row[2]; 
    const clientType = row[3]; 
    const lastDate   = row[4]; 
    
    if (!clientName) return;

    // Create a record for each year
    years.forEach((year, index) => {
      // UPDATED: Shifted start index from 4 to 5 (Column F)
      const colIdx = 5 + (index * 3);
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

// NEW: Ultra-Fast Logic reading from "Dashboard_Cache"
function getDataFromAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Dashboard_Cache");
  
  if (!sheet) return []; // Return empty if missing

  const lastRow = sheet.getLastRow();
  // Adjust startRow to 2 if you have headers in Row 1
  if (lastRow < 2) return [];

  // Fetch Cols A to F (1 to 6)
  // Col A: Client, B: Type, C: Year, D: Qty, E: Orders, F: Max Date
  const rawData = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  
  // Use Script TimeZone for consistent formatting
  const timeZone = ss.getSpreadsheetTimeZone() || "GMT";

  const clientMap = {};

  rawData.forEach(row => {
    const client = row[0];
    const type   = row[1];
    const year   = row[2];
    const qty    = Number(row[3]) || 0;
    const orders = Number(row[4]) || 0;
    const rawDate= row[5]; 

    // Skip invalid rows
    if (!client || !year) return;

    // --- DATE CONVERSION (Strict String) ---
    let dateStr = "";
    if (rawDate instanceof Date) {
      // Convert Date Object -> String "YYYY-MM-DD"
      dateStr = Utilities.formatDate(rawDate, timeZone, "yyyy-MM-dd");
    } else if (typeof rawDate === 'string' && rawDate.length >= 10) {
      // Convert "2023/05/20" -> "2023-05-20"
      dateStr = rawDate.replace(/\//g, "-").slice(0, 10);
    }
    // ---------------------------------------

    if (!clientMap[client]) {
      // Initialize with "0000-00-00" so comparison works
      clientMap[client] = { type: type, lastTx: "0000-00-00", years: {} };
    }

    // Keep the most recent date string
    if (dateStr > clientMap[client].lastTx) {
      clientMap[client].lastTx = dateStr;
    }

    clientMap[client].years[year] = { qty: qty, orders: orders };
  });

  // Flatten Output
  const output = [];
  for (const [name, data] of Object.entries(clientMap)) {
    for (const [year, stats] of Object.entries(data.years)) {
      output.push({
        client: name,
        type: data.type,
        year: year,
        qty: stats.qty,
        orders: stats.orders,
        avg: stats.orders > 0 ? Math.round(stats.qty / stats.orders) : 0,
        
        // CRITICAL FIX: The frontend expects 'lastTx', NOT 'lastDate'
        lastTx: data.lastTx === "0000-00-00" ? "" : data.lastTx
      });
    }
  }
  
  return output;
}
