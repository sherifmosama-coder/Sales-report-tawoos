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
  const rawData = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  
  const timeZone = ss.getSpreadsheetTimeZone() || "GMT";

  // Map rows to Objects (No Aggregation here - we send detailed data)
  const output = rawData.map(row => {
    const rawDate = row[7]; // Adjusted Index for Col H
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
      rev: Number(row[5]) || 0, // New Revenue Field
      orders: Number(row[6]) || 0,
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

// NEW: Fetch Product Data (Hybrid Dynamic Cache)
function getProductData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Product_Cache");
  let liveData = [];
  
  const lastCol = sheet.getLastColumn();
  if (!sheet || sheet.getLastRow() < 2 || lastCol < 2) return { liveData: [], archiveString: "[]" };
  
  // 1. AGGRESSIVE HEADER SCAN (Scans top 3 rows & uses fuzzy text matching)
  const headerBlock = sheet.getRange(1, 2, 3, lastCol - 1).getValues();
  let headers = headerBlock[0];
  for (let i = 0; i < 3; i++) {
     if (headerBlock[i].join('').toLowerCase().includes('qty')) {
         headers = headerBlock[i];
         break;
     }
  }

  const segmentMap = {};
  headers.forEach((h, idx) => {
     let headerStr = String(h).toLowerCase();
     // Fuzzy match: Looks for "qty" and "rev" anywhere inside brackets, ignoring invisible spaces
     if (headerStr.includes('qty') && headerStr.includes('[')) {
        let segName = headerStr.replace(/\[.*?qty.*?\]/i, '').trim();
        segName = segName.replace(/\]|\[/g, '').trim(); // Failsafe for rogue brackets
        if (segName) {
           if (!segmentMap[segName]) segmentMap[segName] = {};
           segmentMap[segName].qtyIdx = idx;
        }
     } else if (headerStr.includes('rev') && headerStr.includes('[')) {
        let segName = headerStr.replace(/\[.*?rev.*?\]/i, '').trim();
        segName = segName.replace(/\]|\[/g, '').trim();
        if (segName) {
           if (!segmentMap[segName]) segmentMap[segName] = {};
           segmentMap[segName].revIdx = idx;
        }
     }
  });

  // 2. Fetch ONLY Live Year (2026) starting strictly from Row 11839
  const startRow = 11839;
  if (sheet.getLastRow() >= startRow) {
    const numRows = sheet.getLastRow() - startRow + 1;
    const rawData = sheet.getRange(startRow, 2, numRows, lastCol - 1).getValues();
    const timeZone = ss.getSpreadsheetTimeZone() || "GMT";
    const currentYear = new Date().getFullYear();

    rawData.forEach(row => {
      let rowYear = Number(row[4]);
      if (rowYear >= currentYear) {
        let rawDate = row[0];
        let dateStr = (rawDate instanceof Date) ? Utilities.formatDate(rawDate, timeZone, "yyyy-MM-dd") : String(rawDate).substring(0, 10);
        
        // Extract dynamically mapped segments
        let rowSegments = {};
        Object.keys(segmentMap).forEach(segName => {
           rowSegments[segName] = {
              qty: Number(row[segmentMap[segName].qtyIdx]) || 0,
              rev: Number(row[segmentMap[segName].revIdx]) || 0
           };
        });

        if (row[1] && dateStr) {
          liveData.push({
            date: dateStr, item: String(row[1]), 
            qty: Number(row[2]) || 0, rev: Number(row[3]) || 0,
            year: String(rowYear), month: String(row[5]), 
            line: String(row[6]), category: String(row[7]),
            segments: rowSegments // Bind the segments mapping
          });
        }
      }
    });
  }

  // 3. Fetch Static Archive (2021-2025) directly as a fast String
  let archiveString = "[]";
  const fileId = PropertiesService.getScriptProperties().getProperty('ARCHIVE_FILE_ID');
  if (fileId) {
     try {
       const file = DriveApp.getFileById(fileId);
       archiveString = file.getBlob().getDataAsString();
     } catch(e) { archiveString = "[]"; }
  }

  return { liveData: liveData, archiveString: archiveString };
}

// NEW: Sync Historical Data to JSON File
function syncHistoricalArchive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Product_Cache");
  const lastCol = sheet.getLastColumn();
  if (!sheet || sheet.getLastRow() < 2 || lastCol < 2) return "No data found.";
  
  // AGGRESSIVE HEADER SCAN (Scans top 3 rows & uses fuzzy text matching)
  const headerBlock = sheet.getRange(1, 2, 3, lastCol - 1).getValues();
  let headers = headerBlock[0];
  for (let i = 0; i < 3; i++) {
     if (headerBlock[i].join('').toLowerCase().includes('qty')) {
         headers = headerBlock[i];
         break;
     }
  }

  const segmentMap = {};
  headers.forEach((h, idx) => {
     let headerStr = String(h).toLowerCase();
     if (headerStr.includes('qty') && headerStr.includes('[')) {
        let segName = headerStr.replace(/\[.*?qty.*?\]/i, '').trim();
        segName = segName.replace(/\]|\[/g, '').trim(); 
        if (segName) {
           if (!segmentMap[segName]) segmentMap[segName] = {};
           segmentMap[segName].qtyIdx = idx;
        }
     } else if (headerStr.includes('rev') && headerStr.includes('[')) {
        let segName = headerStr.replace(/\[.*?rev.*?\]/i, '').trim();
        segName = segName.replace(/\]|\[/g, '').trim();
        if (segName) {
           if (!segmentMap[segName]) segmentMap[segName] = {};
           segmentMap[segName].revIdx = idx;
        }
     }
  });

  // Fetch strictly Rows 2 to 11838 (Years 2021-2025)
  const rawData = sheet.getRange(2, 2, 11837, lastCol - 1).getValues();
  const timeZone = ss.getSpreadsheetTimeZone() || "GMT";
  const currentYear = new Date().getFullYear();
  let archive = [];
  
  rawData.forEach(row => {
    let rowYear = Number(row[4]);
    if (rowYear < currentYear && rowYear > 2000) { 
      let rawDate = row[0];
      let dateStr = (rawDate instanceof Date) ? Utilities.formatDate(rawDate, timeZone, "yyyy-MM-dd") : String(rawDate).substring(0, 10);
      let q = Number(row[2]) || 0;
      let r = Number(row[3]) || 0;
      
      let rowSegments = {};
      Object.keys(segmentMap).forEach(segName => {
         rowSegments[segName] = {
            qty: Number(row[segmentMap[segName].qtyIdx]) || 0,
            rev: Number(row[segmentMap[segName].revIdx]) || 0
         };
      });

      if (row[1] && dateStr && (q !== 0 || r !== 0)) {
         archive.push({
           date: dateStr, item: String(row[1]), 
           qty: q, rev: r, year: String(rowYear), 
           month: String(row[5]), line: String(row[6]), 
           category: String(row[7]), segments: rowSegments
         });
      }
    }
  });

  const fileName = "Tawoos_Product_Archive.json";
  
  // 🔴 IMPORTANT: Paste your personal Folder ID here again!
  const FOLDER_ID = "1nKJ6hIXx200DrywN4AtGFTHugMvdQjuJ"; 
  const folder = DriveApp.getFolderById(FOLDER_ID);
  
  const files = folder.getFilesByName(fileName);
  let file;
  
  if (files.hasNext()) {
    file = files.next();
    file.setContent(JSON.stringify(archive));
  } else {
    file = folder.createFile(fileName, JSON.stringify(archive), MimeType.PLAIN_TEXT);
  }
  
  PropertiesService.getScriptProperties().setProperty('ARCHIVE_FILE_ID', file.getId());
  return "Archive synchronized successfully! " + archive.length + " historical records packaged.";
}

// ==========================================
// NEW PHASE 4: ERP COSTING FETCH
// ==========================================
function getCostingData() {
  try {
      // 1. AUTOSYNC: Run the engine to calculate any new pallets before loading the dashboard
      syncCostLedger();

      // 2. FETCH DATA
      const ss = SpreadsheetApp.openById("1NTLovSrQLtFfebXrSOuWitB29VUV4ifLHmc-Rt_MxWo");
      const sheet = ss.getSheetByName("Items costs");
      
      if (!sheet || sheet.getLastRow() < 2) return [];

      const rawData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
      const timeZone = ss.getSpreadsheetTimeZone() || "GMT";

      let output = [];
      rawData.forEach(row => {
        let rawDate = row[0];
        if (!rawDate) return;
        
        let dateStr = (rawDate instanceof Date) ? Utilities.formatDate(rawDate, timeZone, "yyyy-MM-dd") : String(rawDate).substring(0, 10);
        
        let parsedBomActual = {};
        let parsedBomMarket = {};
        try { if (row[6]) parsedBomActual = JSON.parse(String(row[6])); } catch(e) {}
        try { if (row[10]) parsedBomMarket = JSON.parse(String(row[10])); } catch(e) {}
        
        output.push({
          date: dateStr,
          product: String(row[1]).trim(),
          qty: Number(row[3]) || 0,
          unitActual: Number(row[4]) || 0,
          totalActual: Number(row[5]) || 0,
          bomActual: parsedBomActual,
          planId: String(row[7]),
          unitMarket: Number(row[8]) || 0,
          totalMarket: Number(row[9]) || 0,
          bomMarket: parsedBomMarket
        });
      });
      
      output.sort((a, b) => new Date(b.date) - new Date(a.date));
      return output;
  } catch(e) {
      return { error: e.message };
  }
}

function forceResyncCostLedger() {
  const ss = SpreadsheetApp.openById("1NTLovSrQLtFfebXrSOuWitB29VUV4ifLHmc-Rt_MxWo");
  const sheet = ss.getSheetByName("Items costs");
  
  if (sheet && sheet.getLastRow() > 1) {
    // Clear everything except the headers on Row 1
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }
  
  // Re-run the engine from the beginning of time
  return syncCostLedger();
}
