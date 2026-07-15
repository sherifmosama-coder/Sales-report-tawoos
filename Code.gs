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


function parseNumber(val) {
  if (!val) return 0;
  if (val.toString().includes("#DIV/0!")) return 0;
  const clean = val.toString().replace(/,/g, '');
  const num = parseFloat(clean);
  return isNaN(num) ? 0 : num;
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
  ScriptApp.newTrigger("generateNetSalesCache").timeBased().everyHours(3).create();
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

// Helper to find the absolute latest date in the dataset
function getLatestDataDate(arr, dateKey, isMs) {
  if (!arr || arr.length === 0) return "N/A";
  let max = 0;
  for (let i = 0; i < arr.length; i++) {
    let val = arr[i][dateKey];
    if (!val) continue;
    let ms = isMs ? val : new Date(val).getTime();
    if (ms > max) max = ms;
  }
  if (max === 0) return "N/A";
  const d = new Date(max);
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return `${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear()}`;
}

// --- 1. DASHBOARD DELTA & REBUILD ---
function generateDashboardCache(forceFullRebuild = false) {
  const CUTOFF_MS = new Date().getTime() - (12 * 24 * 60 * 60 * 1000);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Dashboard_Cache"); 
  if (!sheet) return { lastUpdated: new Date().getTime(), latestData: "N/A", data: [] };

  // If Full Rebuild is triggered, act as if there is no existing cache
  let existingCache = forceFullRebuild ? null : readCacheFile("Tawoos_Cache_Dashboard.json");
  let output = (existingCache && existingCache.data) ? existingCache.data.filter(item => {
      return new Date(item.lastTx).getTime() < CUTOFF_MS;
  }) : [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { lastUpdated: new Date().getTime(), latestData: getLatestDataDate(output, 'lastTx', false), data: output };
  
  // Uses 10 Columns to capture Month and Quarter
  const rawData = sheet.getRange(2, 1, lastRow - 1, 10).getValues();

  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    let clientName = row[0]; let yearVal = row[3];
    if (!clientName || !yearVal) continue;

    let rawDate = row[9];
    let dateMs = (rawDate instanceof Date) ? rawDate.getTime() : new Date(rawDate).getTime();
    if (existingCache && existingCache.data && dateMs < CUTOFF_MS) continue;

    let dateStr = "";
    if (rawDate instanceof Date) {
      dateStr = rawDate.getFullYear() + "-" + String(rawDate.getMonth() + 1).padStart(2, '0') + "-" + String(rawDate.getDate()).padStart(2, '0');
    } else if (rawDate) {
      dateStr = String(rawDate).replace(/\//g, "-").slice(0, 10);
    }

    output.push({
      client: clientName, branch: row[1], type: row[2], year: yearVal,
      month: Number(row[4]) + 1, quarter: Number(row[5]),
      qty: Number(row[6]) || 0, rev: Number(row[7]) || 0, orders: Number(row[8]) || 0, lastTx: dateStr
    });
  }

  writeCacheFile("Tawoos_Cache_Dashboard.json", output);
  return { lastUpdated: new Date().getTime(), latestData: getLatestDataDate(output, 'lastTx', false), data: output };
}

// --- 2. CADENCE DELTA & REBUILD ---
function generateCadenceCache(forceFullRebuild = false) {
  const CUTOFF_MS = new Date().getTime() - (12 * 24 * 60 * 60 * 1000);
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

  let existingCache = forceFullRebuild ? null : readCacheFile("Tawoos_Cache_Cadence.json");
  let output = (existingCache && existingCache.data) ? existingCache.data.filter(item => item.date < CUTOFF_MS) : [];

  const masterSheet = ss.getSheetByName("All Data"); 
  if (!masterSheet) return { lastUpdated: new Date().getTime(), latestData: getLatestDataDate(output, 'date', true), data: output };

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
  return { lastUpdated: new Date().getTime(), latestData: getLatestDataDate(output, 'date', true), data: output };
}

// --- 3. PRODUCTS DELTA & REBUILD ---
function generateProductsCache(forceFullRebuild = false) {
  const CUTOFF_MS = new Date().getTime() - (12 * 24 * 60 * 60 * 1000);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let existingCache = forceFullRebuild ? null : readCacheFile("Tawoos_Cache_Products.json");
  
  let combinedData = [];
  if (existingCache && existingCache.data && existingCache.data.length > 0) {
      combinedData = existingCache.data.filter(item => new Date(item.date).getTime() < CUTOFF_MS);
  } else {
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
  return { lastUpdated: new Date().getTime(), latestData: getLatestDataDate(combinedData, 'date', false), data: combinedData };
}

// --- 4. COSTING DELTA & REBUILD ---
function generateCostingCache(forceFullRebuild = false) {
  const CUTOFF_MS = new Date().getTime() - (12 * 24 * 60 * 60 * 1000);
  try {
      syncCostLedger(); 
      
      let existingCache = forceFullRebuild ? null : readCacheFile("Tawoos_Cache_Costing.json");
      let output = (existingCache && existingCache.data) ? existingCache.data.filter(item => {
          return new Date(item.date).getTime() < CUTOFF_MS;
      }) : [];

      const ss = SpreadsheetApp.openById("1NTLovSrQLtFfebXrSOuWitB29VUV4ifLHmc-Rt_MxWo");
      const sheet = ss.getSheetByName("Items costs");
      if (!sheet || sheet.getLastRow() < 2) return { lastUpdated: new Date().getTime(), latestData: "N/A", data: output };

      const rawData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues();

      for (let i = 0; i < rawData.length; i++) {
        const row = rawData[i];
        let rawDate = row[0];
        if (!rawDate) continue;
        
        let dateMs = (rawDate instanceof Date) ? rawDate.getTime() : new Date(rawDate).getTime();
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
      return { lastUpdated: new Date().getTime(), latestData: getLatestDataDate(output, 'date', false), data: output };
  } catch(e) {
      return { error: e.message };
  }
}

// =====================================================================
// ROOT CAUSE SALES PORTAL - JSON ENGINE
// =====================================================================

/**
 * Builds an in-memory dictionary (Hash Map) from the Products_Map sheet.
 * Coding Lesson: Instead of searching the sheet 12,000 times, we read it once.
 * Now, asking for pMap["Premium Tahini"] instantly returns its category.
 */
function buildProductHashMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mapSheet = ss.getSheetByName("Products_Map");
  if (!mapSheet) return {};
  
  // Data starts on row 2, so we ignore the header (index 0)
  const data = mapSheet.getDataRange().getValues();
  let pMap = {};
  
  for (let i = 1; i < data.length; i++) { 
    let pName = String(data[i][0]).trim(); // Col A
    if (!pName) continue;
    
    let isPrivate = String(data[i][1]).toUpperCase() === "TRUE"; // Col B
    let category = String(data[i][2]).trim() || "Uncategorized"; // Col C
    
    pMap[pName] = {
      type: isPrivate ? "Private Label" : "OWN PRODUCT",
      category: category
    };
  }
  
  return pMap;
}

/**
 * Main engine to compress the 12k+ Sales Matrix into a nested JSON file.
 * @param {boolean} isHardRebuild - If true, processes all 12k rows. If false, only the last 12 days.
 */
function generateSalesRootCauseJSON(isHardRebuild = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("All Data");
  const timeZone = ss.getSpreadsheetTimeZone() || "GMT";
  
  // 1. Initialize the Product Dictionary
  const pMap = buildProductHashMap();
  
  // 2. Locate or Initialize the JSON file in Google Drive
  const fileName = "sales_root_cause_cache.json";
  let files = DriveApp.getFilesByName(fileName);
  let file = files.hasNext() ? files.next() : null;
  let jsonData = {};
  
  if (file && !isHardRebuild) {
    try {
      jsonData = JSON.parse(file.getBlob().getDataAsString());
    } catch(e) {
      jsonData = {}; // Reset if corrupted
    }
  }

  // 3. Establish the Delta Update Threshold (12 Days Ago)
  let thresholdMs = 0;
  if (!isHardRebuild) {
    let d = new Date();
    d.setDate(d.getDate() - 12);
    d.setHours(0, 0, 0, 0);
    thresholdMs = d.getTime();
    
    // Coding Lesson: Key Deletion. To update the last 12 days, we first 
    // delete them from the existing JSON. This prevents adding "double sales".
    Object.keys(jsonData).forEach(dateStr => {
      let dMs = new Date(dateStr).getTime();
      if (dMs >= thresholdMs) {
        delete jsonData[dateStr]; 
      }
    });
  }

  // 4. Read the Massive Sales Matrix into Memory
  const lastRow = dataSheet.getLastRow();
  const lastCol = dataSheet.getLastColumn();
  if (lastRow < 10) return "No data found.";
  
  // We grab everything from Row 8 downward.
  // matrixData[0] = Row 8 (Headers), matrixData[1] = Row 9, matrixData[2] = Row 10 (First Data)
  const matrixData = dataSheet.getRange(8, 1, lastRow - 7, lastCol).getValues();
  const headersRow = matrixData[0]; 
  
  // Map out exactly where every product lives dynamically
  // Col R is index 17 in Arrays (A=0, B=1... R=17)
  let productCols = []; 
  for (let c = 17; c < lastCol; c += 2) {
    let pName = String(headersRow[c]).trim();
    if (pName) {
      productCols.push({ name: pName, qtyIdx: c, valIdx: c + 1 });
    }
  }

  // 5. Build the Multi-Dimensional Data Tree
  for (let i = 2; i < matrixData.length; i++) { // Start at index 2 (Row 10)
    let row = matrixData[i];
    let clientName = String(row[2]).trim(); // Col C -> Index 2 (Branch)
    let mainClient = String(row[3]).trim() || clientName; // Col D -> Index 3 (Main Client)
    let segment = String(row[4]).trim() || "Uncategorized"; // Col E -> Index 4
    let dateVal = row[6]; // Col G -> Index 6
    
    if (!dateVal || !clientName) continue;
    
    let dateObj = new Date(dateVal);
    if (isNaN(dateObj.getTime())) continue;
    dateObj.setHours(0, 0, 0, 0);
    let dateMs = dateObj.getTime();
    
    // Skip old rows if we are doing a 12-day soft update
    if (!isHardRebuild && dateMs < thresholdMs) continue;
    
    let dateStr = Utilities.formatDate(dateObj, timeZone, "yyyy-MM-dd");
    
    // Safely build the nested tree path if it doesn't exist yet
    if (!jsonData[dateStr]) jsonData[dateStr] = {};
    if (!jsonData[dateStr][segment]) jsonData[dateStr][segment] = {};
    if (!jsonData[dateStr][segment][mainClient]) jsonData[dateStr][segment][mainClient] = {};
    if (!jsonData[dateStr][segment][mainClient][clientName]) jsonData[dateStr][segment][mainClient][clientName] = {};
    
    let clientNode = jsonData[dateStr][segment][mainClient][clientName];
    
    // Scan all product columns for this specific row
    productCols.forEach(p => {
      let qty = Number(row[p.qtyIdx]) || 0;
      let val = Number(row[p.valIdx]) || 0;
      
      if (qty !== 0 || val !== 0) { // Only log if there is a transaction
        let meta = pMap[p.name] || { category: "Uncategorized", type: "OWN PRODUCT" };
        let cat = meta.category;
        let type = meta.type;
        
        if (!clientNode[cat]) clientNode[cat] = {};
        if (!clientNode[cat][type]) clientNode[cat][type] = {};
        if (!clientNode[cat][type][p.name]) {
          clientNode[cat][type][p.name] = { q: 0, v: 0 }; // Short keys = smaller file size
        }
        
        clientNode[cat][type][p.name].q += qty;
        clientNode[cat][type][p.name].v += val;
      }
    });
  }

  // 6. Overwrite the Google Drive Cache
  const jsonString = JSON.stringify(jsonData);
  if (file) {
    file.setContent(jsonString);
  } else {
    DriveApp.createFile(fileName, jsonString, MimeType.PLAIN_TEXT);
  }
  
  return isHardRebuild ? "Hard Rebuild JSON Cache Complete." : "12-Day Delta Sync Complete.";
}


// =====================================================================
// AUTOMATION TRIGGERS
// =====================================================================

/**
 * Execute this function ONCE manually from the editor to set up the timer.
 */
function setupSalesJSONAutoTrigger() {
  // Clear existing triggers to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "runScheduledDeltaUpdate") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create the 6-hour loop
  ScriptApp.newTrigger("runScheduledDeltaUpdate")
           .timeBased()
           .everyHours(6)
           .create();
}

/**
 * The function fired by the automated timer
 */
function runScheduledDeltaUpdate() {
  generateSalesRootCauseJSON(false); // false = Delta Update
}

/**
 * Call from the WebApp UI for manual rebuilds
 */
function triggerManualSalesRebuild() {
  return generateSalesRootCauseJSON(true); // true = Hard Rebuild
}

/**
 * Called by the frontend to instantly fetch the cached Root Cause JSON
 */
function getRootCauseDataFromCache() {
  let files = DriveApp.getFilesByName("sales_root_cause_cache.json");
  if (!files.hasNext()) return { data: {}, lastUpdated: new Date().toISOString() };
  let file = files.next();
  return {
    data: JSON.parse(file.getBlob().getDataAsString()),
    lastUpdated: file.getLastUpdated().toISOString(),
    latestData: new Date().toISOString()
  };
}
