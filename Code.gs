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

// ==========================================
// MAIN DASHBOARD CACHE GENERATOR
// ==========================================
function getDashboardDataFromCache() {
  let cache = readCacheFile("Tawoos_Cache_Dashboard.json");
  if (!cache) return generateDashboardCache();
  return cache;
}

// ==========================================
// MAIN DASHBOARD CACHE GENERATOR
// ==========================================
function getDashboardDataFromCache() {
  let cache = readCacheFile("Tawoos_Cache_Dashboard.json");
  if (!cache) return generateDashboardCache();
  return cache;
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
  const FOLDER_ID = "1m25aWhRTLHCuyiVgf301I7XE_xc0xIb0"; 
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

      // NOW READING 12 COLUMNS!
      const rawData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues();
      const timeZone = ss.getSpreadsheetTimeZone() || "GMT";

      let output = [];
      rawData.forEach(row => {
        let rawDate = row[0];
        if (!rawDate) return;
        
        let dateStr = (rawDate instanceof Date) ? Utilities.formatDate(rawDate, timeZone, "yyyy-MM-dd") : String(rawDate).substring(0, 10);
        
        let parsedBomActual = {}; let parsedBomMarket = {}; let parsedBomMeta = {};
        try { if (row[6]) parsedBomActual = JSON.parse(String(row[6])); } catch(e) {}
        try { if (row[10]) parsedBomMarket = JSON.parse(String(row[10])); } catch(e) {}
        try { if (row[11]) parsedBomMeta = JSON.parse(String(row[11])); } catch(e) {} // Parse the Meta Column
        
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
          bomMarket: parsedBomMarket,
          bomMeta: parsedBomMeta // Pass the metadata to the UI
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

// ==========================================
// NEW: MANUAL MARKET PRICE MANAGEMENT
// ==========================================
function getManualUpdates() {
   const ss = SpreadsheetApp.openById("1NTLovSrQLtFfebXrSOuWitB29VUV4ifLHmc-Rt_MxWo");
   let sheet = ss.getSheetByName("Price Updates");
   if(!sheet) return [];
   
   let data = sheet.getDataRange().getValues();
   let res = [];
   let timeZone = ss.getSpreadsheetTimeZone() || "GMT";
   
   for(let i=1; i<data.length; i++) {
       if(data[i][0] && data[i][1]) {
           let d = data[i][0];
           let dStr = (d instanceof Date) ? Utilities.formatDate(d, timeZone, "yyyy-MM-dd") : String(d).substring(0,10);
           res.push({date: dStr, material: String(data[i][1]).trim(), price: Number(data[i][2]).toFixed(2)});
       }
   }
   return res.reverse(); // Return newest first for the table
}

function saveAndResyncMarketPrice(dateStr, material, price) {
   const ss = SpreadsheetApp.openById("1NTLovSrQLtFfebXrSOuWitB29VUV4ifLHmc-Rt_MxWo");
   let sheet = ss.getSheetByName("Price Updates");
   
   if(!sheet) {
      sheet = ss.insertSheet("Price Updates");
      sheet.appendRow(["Date", "Material", "Price"]);
   }
   
   // Insert new update
   sheet.appendRow([dateStr, material, Number(price)]);
   
   // Immediately trigger global Cost Engine re-sync so dashboard updates instantly
   forceResyncCostLedger(); 
   return true;
}

// ==========================================
// NEW: CLIENT ORDER CADENCE FETCH (SMART TWO-TIER + CONTACTS)
// ==========================================
// ==========================================
// CADENCE CACHE GENERATOR
// ==========================================
function getCadenceDataFromCache() {
  let cache = readCacheFile("Tawoos_Cache_Cadence.json");
  // If the file doesn't exist yet, force generate it right now
  if (!cache) return generateCadenceCache();
  return cache;
}


// ==========================================
// NEW: CADENCE EXPORT ENGINE (PDF & EXCEL)
// ==========================================
function generateCadenceExportUrl(payloadStr, format) {
  const data = JSON.parse(payloadStr);
  const ss = SpreadsheetApp.create("[TEMP EXPORT] Late Clients Watchlist");
  const sheet = ss.getSheets()[0];

  // 1. Format Headers
  const headers = ["Client Name", "Segment", "Region", "Contact Data", "Category", "Last Order Qty", "Last Order Date", "Expected Cycle (Days)", "Days Late"];
  sheet.appendRow(headers);
  let headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold").setBackground("#3b82f6").setFontColor("white");

  let rowIndex = 2;

  // 2. Populate Data with Grouping
  data.forEach(item => {
    let orderDate = new Date(item.lastOrderDate).toLocaleDateString('en-GB');

    // Main Row (Summary View)
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([[
      item.client, item.segment, item.region, item.contact, item.category,
      item.lastOrderQty, orderDate, item.expectedInterval, item.daysLate
    ]]);
    sheet.getRange(rowIndex, 1, 1, headers.length).setFontWeight("bold").setBackground("#f8f9fa");
    rowIndex++;

    let startGroupRow = rowIndex;

    // Detail Rows (Order History)
    if (item.history && item.history.length > 0) {
       item.history.forEach(h => {
          let hDate = new Date(h.date).toLocaleDateString('en-GB');
          let details = h.items.map(i => `${i.qty}x ${i.name}`).join(' | ');

          sheet.getRange(rowIndex, 1, 1, headers.length).setValues([[
            "↳ Order Details", "", "", details, "", h.totalQty, hDate, "", ""
          ]]);
          sheet.getRange(rowIndex, 1, 1, headers.length).setFontColor("#6c757d");
          rowIndex++;
       });

       // Excel Grouping: Shift these detail rows into a collapsible group
       let groupRange = sheet.getRange(startGroupRow, 1, item.history.length, sheet.getMaxColumns());
       groupRange.shiftRowGroupDepth(1);
    }
  });

  // 3. Final Beautification
  sheet.collapseAllRowGroups(); // Collapses everything to show only the Main Rows initially
  
  // Format rows: Align everything to the top so it looks clean when wrapped
  sheet.getRange(1, 1, rowIndex, headers.length).setVerticalAlignment("top");
  
  // Wrap text in the Contact Data column (Col 4) so phone numbers sit on multiple lines
  sheet.getRange(1, 4, rowIndex, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  
  // Auto-resize all columns to perfectly fit their contents
  sheet.autoResizeColumns(1, headers.length);
  
  // Safety net: ensure the Contact column isn't squeezed too tightly after autofit
  if (sheet.getColumnWidth(4) < 200) {
      sheet.setColumnWidth(4, 250); 
  }
  
  SpreadsheetApp.flush();

  // 4. Generate direct download URL
  let url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=${format}&portrait=false&fitw=true`;

  return { url: url, fileId: ss.getId() };
}

function deleteTempCadenceFile(fileId) {
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
  } catch (e) {}
}


// ==========================================
// NEW: MASTER CACHE ENGINE & TRIGGERS
// ==========================================
// We use your existing Folder ID to instantly bypass Google Drive's slow search algorithm!
const CACHE_FOLDER_ID = "1m25aWhRTLHCuyiVgf301I7XE_xc0xIb0";

function getCacheFolder() {
  return DriveApp.getFolderById(CACHE_FOLDER_ID);
}

function writeCacheFile(fileName, dataObj) {
  const folder = getCacheFolder();
  let files = folder.getFilesByName(fileName);
  
  // TRASH OLD FILES FIRST (Prevents setContent hangs and duplicates)
  while (files.hasNext()) {
    files.next().setTrashed(true);
  }
  
  // CREATE FRESH FILE
  let payload = JSON.stringify({ lastUpdated: new Date().getTime(), data: dataObj });
  folder.createFile(fileName, payload, MimeType.PLAIN_TEXT);
}
function readCacheFile(fileName) {
  const folder = getCacheFolder();
  let files = folder.getFilesByName(fileName);
  if (files.hasNext()) {
    return JSON.parse(files.next().getBlob().getDataAsString());
  }
  return null;
}

function setupCacheTriggers() {
  let triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction().includes("generate")) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Setup ALL robots to run every 3 hours in the background
  ScriptApp.newTrigger("generateDashboardCache").timeBased().everyHours(3).create();
  ScriptApp.newTrigger("generateCadenceCache").timeBased().everyHours(3).create();
  ScriptApp.newTrigger("generateProductsCache").timeBased().everyHours(3).create();
  ScriptApp.newTrigger("generateCostingCache").timeBased().everyHours(3).create();
}

// ==========================================
// PRODUCTS CACHE GENERATOR (Archive + Live Merge)
// ==========================================
function getProductDataFromCache() {
  let cache = readCacheFile("Tawoos_Cache_Products.json");
  if (!cache) return generateProductsCache();
  return cache;
}

// ==========================================
// COSTING CACHE GENERATOR
// ==========================================
function getCostingDataFromCache() {
  let cache = readCacheFile("Tawoos_Cache_Costing.json");
  if (!cache) return generateCostingCache();
  return cache;
}

// ==========================================
// 12-DAY DELTA GENERATORS (V8 OPTIMIZED)
// ==========================================
const DELTA_DAYS = 12;
const CUTOFF_MS = new Date().getTime() - (DELTA_DAYS * 24 * 60 * 60 * 1000);

// --- 1. DASHBOARD DELTA ---
function generateDashboardCache() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Dashboard_Cache"); 
  if (!sheet) return { lastUpdated: new Date().getTime(), data: [] };

  let existingCache = readCacheFile("Tawoos_Cache_Dashboard.json");
  let output = (existingCache && existingCache.data) ? existingCache.data.filter(item => {
      return new Date(item.lastTx).getTime() < CUTOFF_MS;
  }) : [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { lastUpdated: new Date().getTime(), data: output };
  
  const rawData = sheet.getRange(2, 1, lastRow - 1, 8).getValues();

  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    let clientName = row[0]; let yearVal = row[3];
    if (!clientName || !yearVal) continue;

    let rawDate = row[7];
    let dateMs = (rawDate instanceof Date) ? rawDate.getTime() : new Date(rawDate).getTime();
    
    // DELTA FILTER: Skip rows older than 12 days IF we already have a cache
    if (existingCache && existingCache.data && dateMs < CUTOFF_MS) continue;

    let dateStr = "";
    if (rawDate instanceof Date) {
      dateStr = rawDate.getFullYear() + "-" + String(rawDate.getMonth() + 1).padStart(2, '0') + "-" + String(rawDate.getDate()).padStart(2, '0');
    } else if (rawDate) {
      dateStr = String(rawDate).replace(/\//g, "-").slice(0, 10);
    }

    output.push({
      client: clientName, branch: row[1], type: row[2], year: yearVal,
      qty: Number(row[4]) || 0, rev: Number(row[5]) || 0, orders: Number(row[6]) || 0, lastTx: dateStr
    });
  }

  writeCacheFile("Tawoos_Cache_Dashboard.json", output);
  return { lastUpdated: new Date().getTime(), data: output };
}

// --- 2. CADENCE DELTA ---
function generateCadenceCache() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clientsSheet = ss.getSheetByName("Clients");
  let contactMap = {};
  if (clientsSheet) {
    const cData = clientsSheet.getDataRange().getValues();
    for (let i = 1; i < cData.length; i++) {
      let cName = String(cData[i][0]).trim();
      if (cName) contactMap[cName] = { region: String(cData[i][3]).trim() || 'Unspecified', contact: String(cData[i][4]).trim() || 'No contact provided' };
    }
  }

  let existingCache = readCacheFile("Tawoos_Cache_Cadence.json");
  let output = (existingCache && existingCache.data) ? existingCache.data.filter(item => item.date < CUTOFF_MS) : [];

  const masterSheet = ss.getSheetByName("All Data"); 
  if (!masterSheet) return { lastUpdated: new Date().getTime(), data: output };

  const data = masterSheet.getDataRange().getValues();
  const ROW_PROD_NAMES = 7, ROW_DATA_START = 9;
  const COL_CLIENT = 2, COL_SEGMENT = 4, COL_DATE = 6, COL_PROD_START = 17;

  let targetProductCols = {}; 
  let numCols = data[ROW_PROD_NAMES].length;
  for (let c = COL_PROD_START; c < numCols; c += 2) { 
    let itemName = String(data[ROW_PROD_NAMES][c]).trim();
    if (itemName) targetProductCols[c] = itemName;
  }

  for (let i = ROW_DATA_START; i < data.length; i++) {
    let rowDate = data[i][COL_DATE];
    if (!rowDate) continue; 
    let dateMs = (rowDate instanceof Date) ? rowDate.getTime() : new Date(rowDate).getTime();
    if (!dateMs || isNaN(dateMs)) continue;
    
    // DELTA FILTER
    if (existingCache && existingCache.data && dateMs < CUTOFF_MS) continue;

    let clientName = String(data[i][COL_CLIENT]).trim();
    if (!clientName) continue;

    let segmentCache = String(data[i][COL_SEGMENT]).trim() || 'Uncategorized';
    let cInfo = contactMap[clientName] || { region: 'Unspecified', contact: 'No contact provided' };

    for (let colStr in targetProductCols) {
      let c = parseInt(colStr);
      let rawQty = data[i][c];
      if (!rawQty) continue; 
      let qty = Number(rawQty);
      if (qty > 0) {
        output.push({
          date: dateMs, client: clientName, segment: segmentCache,
          region: cInfo.region, contactInfo: cInfo.contact, item: targetProductCols[c], qty: qty
        });
      }
    }
  }

  writeCacheFile("Tawoos_Cache_Cadence.json", output);
  return { lastUpdated: new Date().getTime(), data: output };
}

// --- 3. PRODUCTS DELTA ---
function generateProductsCache() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let existingCache = readCacheFile("Tawoos_Cache_Products.json");
  
  let combinedData = [];
  if (existingCache && existingCache.data && existingCache.data.length > 0) {
      // Keep everything older than 12 days (this automatically preserves the 2021-2025 archive!)
      combinedData = existingCache.data.filter(item => new Date(item.date).getTime() < CUTOFF_MS);
  } else {
      // First run only: fetch archive manually
      const fileId = PropertiesService.getScriptProperties().getProperty('ARCHIVE_FILE_ID');
      if (fileId) {
         try {
           const file = DriveApp.getFileById(fileId);
           combinedData = JSON.parse(file.getBlob().getDataAsString());
         } catch(e) {}
      }
  }

  const sheet = ss.getSheetByName("Product_Cache");
  const lastCol = sheet ? sheet.getLastColumn() : 0;
  if (sheet && sheet.getLastRow() >= 2 && lastCol >= 2) {
      const headerBlock = sheet.getRange(1, 2, 3, lastCol - 1).getValues();
      let headers = headerBlock[0];
      for (let i = 0; i < 3; i++) { if (headerBlock[i].join('').toLowerCase().includes('qty')) { headers = headerBlock[i]; break; } }

      const segmentMap = {};
      headers.forEach((h, idx) => {
         let headerStr = String(h).toLowerCase();
         if (headerStr.includes('qty') && headerStr.includes('[')) {
            let segName = headerStr.replace(/\[.*?qty.*?\]/i, '').replace(/\]|\[/g, '').trim();
            if (segName) { if (!segmentMap[segName]) segmentMap[segName] = {}; segmentMap[segName].qtyIdx = idx; }
         } else if (headerStr.includes('rev') && headerStr.includes('[')) {
            let segName = headerStr.replace(/\[.*?rev.*?\]/i, '').replace(/\]|\[/g, '').trim();
            if (segName) { if (!segmentMap[segName]) segmentMap[segName] = {}; segmentMap[segName].revIdx = idx; }
         }
      });

      const startRow = 11839;
      if (sheet.getLastRow() >= startRow) {
        const numRows = sheet.getLastRow() - startRow + 1;
        const rawData = sheet.getRange(startRow, 2, numRows, lastCol - 1).getValues();
        const currentYear = new Date().getFullYear();
        
        for (let i = 0; i < rawData.length; i++) {
          const row = rawData[i];
          let rowYear = Number(row[4]);
          
          if (rowYear >= currentYear && row[1]) {
            let rawDate = row[0];
            let dateMs = (rawDate instanceof Date) ? rawDate.getTime() : new Date(rawDate).getTime();
            
            // DELTA FILTER
            if (existingCache && existingCache.data && dateMs < CUTOFF_MS) continue;

            let dateStr = "";
            if (rawDate instanceof Date) {
              dateStr = rawDate.getFullYear() + "-" + String(rawDate.getMonth() + 1).padStart(2, '0') + "-" + String(rawDate.getDate()).padStart(2, '0');
            } else if (rawDate) {
              dateStr = String(rawDate).substring(0, 10);
            }
            
            let rowSegments = {};
            Object.keys(segmentMap).forEach(segName => {
               rowSegments[segName] = { qty: Number(row[segmentMap[segName].qtyIdx]) || 0, rev: Number(row[segmentMap[segName].revIdx]) || 0 };
            });

            combinedData.push({
              date: dateStr, item: String(row[1]), qty: Number(row[2]) || 0, rev: Number(row[3]) || 0,
              year: String(rowYear), month: String(row[5]), line: String(row[6]), category: String(row[7]), segments: rowSegments
            });
          }
        }
      }
  }

  writeCacheFile("Tawoos_Cache_Products.json", combinedData);
  return { lastUpdated: new Date().getTime(), data: combinedData };
}

// --- 4. COSTING DELTA ---
function generateCostingCache() {
  try {
      syncCostLedger(); 
      
      let existingCache = readCacheFile("Tawoos_Cache_Costing.json");
      let output = (existingCache && existingCache.data) ? existingCache.data.filter(item => {
          return new Date(item.date).getTime() < CUTOFF_MS;
      }) : [];

      const ss = SpreadsheetApp.openById("1NTLovSrQLtFfebXrSOuWitB29VUV4ifLHmc-Rt_MxWo");
      const sheet = ss.getSheetByName("Items costs");
      if (!sheet || sheet.getLastRow() < 2) return { lastUpdated: new Date().getTime(), data: output };

      const rawData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues();

      for (let i = 0; i < rawData.length; i++) {
        const row = rawData[i];
        let rawDate = row[0];
        if (!rawDate) continue;
        
        let dateMs = (rawDate instanceof Date) ? rawDate.getTime() : new Date(rawDate).getTime();
        
        // DELTA FILTER
        if (existingCache && existingCache.data && dateMs < CUTOFF_MS) continue;

        let dateStr = "";
        if (rawDate instanceof Date) {
          dateStr = rawDate.getFullYear() + "-" + String(rawDate.getMonth() + 1).padStart(2, '0') + "-" + String(rawDate.getDate()).padStart(2, '0');
        } else {
          dateStr = String(rawDate).substring(0, 10);
        }
        
        let parsedBomActual = {}, parsedBomMarket = {}, parsedBomMeta = {};
        if (row[6]) try { parsedBomActual = JSON.parse(String(row[6])); } catch(e) {}
        if (row[10]) try { parsedBomMarket = JSON.parse(String(row[10])); } catch(e) {}
        if (row[11]) try { parsedBomMeta = JSON.parse(String(row[11])); } catch(e) {}
        
        output.push({
          date: dateStr, product: String(row[1]).trim(), qty: Number(row[3]) || 0,
          unitActual: Number(row[4]) || 0, totalActual: Number(row[5]) || 0, bomActual: parsedBomActual,
          planId: String(row[7]), unitMarket: Number(row[8]) || 0, totalMarket: Number(row[9]) || 0, 
          bomMarket: parsedBomMarket, bomMeta: parsedBomMeta
        });
      }
      
      output.sort((a, b) => new Date(b.date) - new Date(a.date));
      writeCacheFile("Tawoos_Cache_Costing.json", output);
      return { lastUpdated: new Date().getTime(), data: output };
  } catch(e) {
      return { error: e.message };
  }
}
