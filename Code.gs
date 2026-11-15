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
  
  // Re-run the calculation engine and overwrite Google Drive JSON Cache
  return generateCostingCache(true);
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
  ScriptApp.newTrigger("generateBranchNetSalesCache").timeBased().everyHours(3).create();
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
  try {
      const ss = SpreadsheetApp.openById("1NTLovSrQLtFfebXrSOuWitB29VUV4ifLHmc-Rt_MxWo");
      syncCostLedger(ss); 
      
      let output = [];
      const sheet = ss.getSheetByName("Items costs");
      if (!sheet || sheet.getLastRow() < 2) return { lastUpdated: new Date().getTime(), latestData: "N/A", data: output };

      const rawData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues();

      for (let i = 0; i < rawData.length; i++) {
        const row = rawData[i];
        let rawDate = row[0];
        if (!rawDate) continue;

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
          planId: String(row[7]), recipeId: (parsedBomMeta && parsedBomMeta._recipeId) ? String(parsedBomMeta._recipeId) : "",
          unitMarket: Number(row[8]) || 0, totalMarket: Number(row[9]) || 0, 
          bomMarket: parsedBomMarket, bomMeta: parsedBomMeta
        });
      }
      
      output.sort((a, b) => new Date(b.date) - new Date(a.date));
      writeCacheFile("Tawoos_Cache_Costing.json", output);
      
      // Clear the Profitability Cache so it forces a fresh rebuild upon next load
      writeCacheFile("Tawoos_Cache_ProfitabilityMerged.json", {});

      return { lastUpdated: new Date().getTime(), latestData: getLatestDataDate(output, 'date', false), data: output };
  } catch(e) {
      return { error: e.message };
  }
}


// ==========================================
// DEDICATED BRANCH-LEVEL NET SALES CACHE GENERATOR
// ==========================================
function getBranchNetSalesDataFromCache() {
  let cache = readCacheFile("Tawoos_Cache_BranchNetSales.json");
  if (!cache) return generateBranchNetSalesCache();
  return cache;
}

function generateBranchNetSalesCache(forceFullRebuild = false) {
  const CUTOFF_MS = new Date().getTime() - (12 * 24 * 60 * 60 * 1000);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("All Data");
  
  let existingCache = forceFullRebuild ? null : readCacheFile("Tawoos_Cache_BranchNetSales.json");
  let output = (existingCache && existingCache.data) ? existingCache.data.filter(item => item.dateMs < CUTOFF_MS) : [];

  if (!masterSheet) return { lastUpdated: new Date().getTime(), latestData: getLatestDataDate(output, 'dateMs', true), data: output };

  const data = masterSheet.getDataRange().getValues();
  const ROW_PROD_NAMES = 7, ROW_DATA_START = 9;
  const COL_BRANCH = 2, COL_MAIN_CLIENT = 3, COL_SEGMENT = 4, COL_DATE = 6, COL_PROD_START = 17;

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

    let branchName = String(data[i][COL_BRANCH]).trim(); // Col C -> Branch Level
    if (!branchName) continue;

    let mainClient = String(data[i][COL_MAIN_CLIENT]).trim() || branchName; // Col D -> Parent Level
    let segmentCache = String(data[i][COL_SEGMENT]).trim() || 'Unassigned';

    for (let colStr in targetProductCols) {
      let c = parseInt(colStr);
      let rawQty = data[i][c];
      let rawVal = data[i][c + 1];
      
      let qty = Number(rawQty) || 0;
      let netRev = Number(rawVal) || 0;

      if (qty !== 0 || netRev !== 0) {
        output.push({
          dateMs: dateMs,
          client: branchName,        // Branch Name (e.g., "هايبر وان الشيخ زايد")
          mainClient: mainClient,    // Main Account (e.g., "هايبر وان")
          segment: segmentCache,
          item: targetProductCols[c], // Product Name
          qty: qty,
          netRev: netRev
        });
      }
    }
  }

  writeCacheFile("Tawoos_Cache_BranchNetSales.json", output);
  return { lastUpdated: new Date().getTime(), latestData: getLatestDataDate(output, 'dateMs', true), data: output };
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

// ==========================================
// DYNAMIC CONFIGURATION & PERSISTENCE
// ==========================================

// ==========================================
// HELPER: SANITIZE & NORMALIZE TIME TO HH:mm
// ==========================================
function formatTimeString(val) {
  if (!val) return '08:00';
  
  // Case 1: Google Sheets auto-cast the cell to a Date object
  if (val instanceof Date) {
    var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || "GMT";
    return Utilities.formatDate(val, tz, "HH:mm");
  }
  
  var str = String(val).trim();
  
  // Case 2: String is formatted as H:mm or HH:mm (e.g., "8:00" -> "08:00")
  if (/^\d{1,2}:\d{2}$/.test(str)) {
    var parts = str.split(':');
    return parts[0].padStart(2, '0') + ':' + parts[1];
  }
  
  return '08:00';
}

// ==========================================
// EMAIL ENGINE CACHE SYNC & TIMESTAMP HELPERS
// ==========================================
function getLastEngineSyncTime() {
  try {
    var lastSync = PropertiesService.getScriptProperties().getProperty("LAST_ENGINE_SYNC_TIME");
    return lastSync || "--";
  } catch (e) {
    return "--";
  }
}

function syncEmailEngineCaches() {
  try {
    // Re-hydrate primary data caches used by email engine sections
    if (typeof generateDashboardCache === 'function') generateDashboardCache(true);
    if (typeof generateBranchNetSalesCache === 'function') generateBranchNetSalesCache(true);
    if (typeof generateCadenceCache === 'function') generateCadenceCache(true);

    var tz = Session.getScriptTimeZone() || "GMT";
    var formattedTime = Utilities.formatDate(new Date(), tz, "dd/MM hh:mm a");
    PropertiesService.getScriptProperties().setProperty("LAST_ENGINE_SYNC_TIME", formattedTime);

    return formattedTime;
  } catch (err) {
    Logger.log("Error in syncEmailEngineCaches: " + err.message);
    throw new Error("Failed to sync engine caches: " + err.message);
  }
}

function getEmailConfigs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Email_Configs");
  
  if (!sheet) {
    sheet = ss.insertSheet("Email_Configs");
    sheet.appendRow([
      "Template_ID", "Template_Name", "Time_Scope", "Enabled", 
      "To_Recipients", "CC_Recipients", "Schedule_Day", "Schedule_Time", 
      "Base_Year", "Measure", "Segments_Included", "Sections_Included", "Last_Sent", "Report_Portal", "Sender_Name", "Language"
    ]);
  }

  var data = sheet.getDataRange().getValues();
  var templates = [];

  if (data.length > 1) {
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[0]) {
        templates.push({
          templateId: String(row[0]),
          templateName: String(row[1] || 'Untitled Alert'),
          timeScope: String(row[2] || 'Daily'),
          enabled: String(row[3]).toUpperCase() === "TRUE",
          toRecipients: String(row[4] || ''),
          ccRecipients: String(row[5] || ''),
          scheduleDay: String(row[6] || 'MON,TUE,WED,THU,FRI'),
          scheduleTime: formatTimeString(row[7]),
          baseYear: String(row[8] || '2025'),
          measure: String(row[9] || 'BOTH'),
          segments: String(row[10] || 'ALL').split(','),
          sections: String(row[11] || '1,2,3,4,5,6,7,8').split(',').map(Number).filter(function(n){ return !isNaN(n); }),
          lastSent: String(row[12] || ''),
          reportPortal: String(row[13] || 'NONE'),
          senderName: String(row[14] || 'Sales Performance Engine'),
          language: String(row[15] || 'AR').trim().toUpperCase() // Column 16 (Language)
        });
      }
    }
  }
  
  return { templates: templates };
}

function saveEmailConfigs(templatesList) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Email_Configs");
  if (!sheet) {
    getEmailConfigs();
    sheet = ss.getSheetByName("Email_Configs");
  }

  // Clear data rows
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 16).clearContent();
  }

  var newRows = [];
  templatesList.forEach(function(t) {
    var segStr = Array.isArray(t.segments) ? t.segments.join(",") : "ALL";
    var secStr = Array.isArray(t.sections) ? t.sections.join(",") : "1,2,3,4,5,6,7,8";
    newRows.push([
      t.templateId,
      t.templateName || "Untitled Alert",
      t.timeScope || "Daily",
      t.enabled ? "TRUE" : "FALSE",
      t.toRecipients || "",
      t.ccRecipients || "",
      t.scheduleDay || "",
      t.scheduleTime || "08:00",
      t.baseYear || "2025",
      t.measure || "BOTH",
      segStr,
      secStr,
      t.lastSent || "",
      t.reportPortal || "NONE",
      t.senderName || "Sales Performance Engine",
      t.language || "AR" // Column 16
    ]);
  });

  if (newRows.length > 0) {
    sheet.getRange(2, 1, newRows.length, 16).setValues(newRows);
  }

  // Update Single Master Runner Trigger
  setupMasterEmailTrigger();
  
  return { success: true };
}

function setupMasterEmailTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var exists = false;
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === "masterScheduledEmailRunner") {
      exists = true;
    }
  });

  if (!exists) {
    // Run master engine daily at 8:00 AM to process applicable schedules
    ScriptApp.newTrigger("masterScheduledEmailRunner")
      .timeBased()
      .everyDays(1)
      .atHour(8)
      .create();
  }
}

// ==========================================
// HEADLESS MASTER EXECUTOR
// ==========================================
function masterScheduledEmailRunner() {
  var res = getEmailConfigs();
  var templates = res.templates;
  if (!templates || templates.length === 0) return;

  var now = new Date();
  var dayMap = ['SUN', 'MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT'];
  var todayStr = dayMap[now.getDay()];
  var dateNum = now.getDate();
  var currentHour = now.getHours(); // 0 to 23

  // Month-end calculation helper
  var tomorrow = new Date(now.getFullYear(), now.getMonth(), dateNum + 1);
  var isLastDayOfMonth = (tomorrow.getDate() === 1);

  var monthIdx = now.getMonth(); // 0-11
  var isQuarterStartMonth = [0, 3, 6, 9].indexOf(monthIdx) !== -1;  // Jan, Apr, Jul, Oct
  var isQuarterEndMonth = [2, 5, 8, 11].indexOf(monthIdx) !== -1;   // Mar, Jun, Sep, Dec

  var currentHourPrefix = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH");

  templates.forEach(function(t) {
    if (!t.enabled || !t.toRecipients) return;

    // 1. HOUR MATCHING: Check if current hour matches scheduleTime (e.g., "08:00" -> 8)
    var tTime = t.scheduleTime || '08:00';
    var tHour = parseInt(tTime.split(':')[0], 10) || 8;
    
    if (currentHour !== tHour) return; // Skip if it's not the designated dispatch hour

    // 2. DEDUPLICATION GUARD: Skip if already sent during this hour
    if (t.lastSent && String(t.lastSent).indexOf(currentHourPrefix) === 0) {
      Logger.log("Template " + t.templateName + " already sent this hour. Skipping.");
      return;
    }

    // 2. DATE & SCOPE MATCHING
    var shouldRun = false;

    if (t.timeScope === 'Daily') {
      var days = (t.scheduleDay || '').split(',');
      if (days.indexOf(todayStr) !== -1) shouldRun = true;

    } else if (t.timeScope === 'Weekly') {
      if (t.scheduleDay === todayStr) shouldRun = true;

    } else if (t.timeScope === 'Monthly') {
      if (t.scheduleDay === 'LAST_DAY') {
        if (isLastDayOfMonth) shouldRun = true;
      } else if (String(t.scheduleDay) === String(dateNum)) {
        shouldRun = true;
      }

    } else if (t.timeScope === 'Quarterly') {
      var qDay = t.scheduleDay || 'Q_START';
      if (qDay === 'Q_END') {
        // Last day of Q1 (Mar 31), Q2 (Jun 30), Q3 (Sep 30), Q4 (Dec 31)
        if (isQuarterEndMonth && isLastDayOfMonth) shouldRun = true;
      } else {
        // First day of Q1 (Jan 1), Q2 (Apr 1), Q3 (Jul 1), Q4 (Oct 1)
        if (isQuarterStartMonth && dateNum === 1) shouldRun = true;
      }
    }

    // 3. EXECUTE DISPATCH IF MATCHED
    if (shouldRun) {
      sendTestEmailDigest(t.templateId, null, t);
    }
  });
}

// ==========================================
// SEGMENT-FILTERED EMAIL GENERATOR
// ==========================================
// ==========================================
// HELPER: STANDARDIZED K/M NUMBER FORMATTER
// ==========================================
function formatKM(val, unit) {
  if (val === null || val === undefined || isNaN(val) || val === 0) return '0' + (unit ? ' ' + unit : '');
  var absVal = Math.abs(val);
  var sign = val < 0 ? '-' : '';
  var formatted = '';
  
  if (absVal >= 1000000) {
    formatted = (absVal / 1000000).toFixed(1).replace(/\.0$/, '') + 'M';
  } else if (absVal >= 1000) {
    formatted = (absVal / 1000).toFixed(1).replace(/\.0$/, '') + 'K';
  } else {
    formatted = absVal.toFixed(1).replace(/\.0$/, '');
  }
  return sign + formatted + (unit ? ' ' + unit : '');
}

// ==========================================
// TIME-SCOPED EMAIL GENERATOR & DISPATCHER
// ==========================================

// ==========================================
// HELPER: DUAL-AXIS ORDER CONSISTENCY CALCULATOR
// ==========================================
function calculateOrderConsistency(currentOrderMs, currentQty, priorOrders, msPerDay) {
  if (!priorOrders || priorOrders.length === 0) {
    return {
      text: "First Order",
      style: "background-color:#f1f5f9; color:#475569;"
    };
  }

  // Slice to max last 4 prior orders
  var recentPrior = priorOrders.slice(-4);
  
  // 1. VOLUME CONSISTENCY
  var sumQty = 0;
  recentPrior.forEach(function(o) { sumQty += o.totalQty; });
  var avgQty = sumQty / recentPrior.length;
  var qtyVarPct = avgQty > 0 ? ((currentQty - avgQty) / avgQty) * 100 : 0;

  var volText = "Normal Vol";
  if (qtyVarPct > 15) volText = "+" + Math.round(qtyVarPct) + "% Vol";
  else if (qtyVarPct < -15) volText = Math.round(qtyVarPct) + "% Vol";

  // 2. TIMING CONSISTENCY
  var lastPrior = recentPrior[recentPrior.length - 1];
  var daysSinceLast = (currentOrderMs - lastPrior.date) / msPerDay;
  
  var expectedCycle = 30; // Fallback default
  if (recentPrior.length >= 2) {
    var firstPrior = recentPrior[0];
    var totalElapsed = (lastPrior.date - firstPrior.date) / msPerDay;
    var qConsumed = 0;
    for (var i = 0; i < recentPrior.length - 1; i++) {
      qConsumed += recentPrior[i].totalQty;
    }
    if (totalElapsed > 0 && qConsumed > 0) {
      var dailyRate = qConsumed / totalElapsed;
      if (dailyRate > 0) {
        expectedCycle = lastPrior.totalQty / dailyRate;
      }
    }
  }
  expectedCycle = Math.max(3, Math.min(180, expectedCycle));

  var gapDiff = daysSinceLast - expectedCycle;
  var timingText = "On Schedule";
  if (gapDiff > 2) timingText = "Delayed (+" + Math.floor(gapDiff) + "d)";
  else if (gapDiff < -2) timingText = "Early (" + Math.ceil(gapDiff) + "d)";

  // 3. BADGE STYLE DETERMINATION
  var badgeStyle = "background-color:#dcfce7; color:#15803d;"; // Green default (Normal/On Schedule)
  if (gapDiff > 5 || qtyVarPct < -30) {
    badgeStyle = "background-color:#fee2e2; color:#b91c1c;"; // Red (Significantly Delayed or Low Vol)
  } else if (gapDiff > 2 || gapDiff < -5 || qtyVarPct > 30) {
    badgeStyle = "background-color:#fef3c7; color:#b45309;"; // Amber (Slightly Delayed / High Vol)
  }

  return {
    text: timingText + " | " + volText,
    style: badgeStyle
  };
}

// ==========================================
// HELPER: UPDATE LAST SENT TIMESTAMP IN SHEET
// ==========================================
function updateTemplateLastSent(templateId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Email_Configs");
    if (!sheet) return;

    var data = sheet.getDataRange().getValues();
    var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(templateId)) {
        sheet.getRange(i + 1, 13).setValue(nowStr); // Column 13 (Last_Sent)
        break;
      }
    }
  } catch (err) {
    Logger.log("Error updating Last_Sent: " + err.message);
  }
}

function sendTestEmailDigest(templateId, recipientOverride, tplObject) {
  var tpl = tplObject;
  if (!tpl) {
    var res = getEmailConfigs();
    tpl = res.templates.find(function(x) { return x.templateId === templateId; });
  }
  if (!tpl) throw new Error("Template settings not found");

  var to = recipientOverride || tpl.toRecipients;
  if (!to) throw new Error("No recipient email specified");

  // Language & Directional Setup
  var lang = (tpl.language || 'AR').toUpperCase();
  var isAr = (lang === 'AR');
  var dir = isAr ? 'rtl' : 'ltr';
  var align = isAr ? 'right' : 'left';
  var revUnit = isAr ? 'ج.م' : 'EGP';
  var qtyUnit = isAr ? 'وحدة' : 'Units';

  // Approved Internationalization Dictionary
  var i18n = {
    scopeLabel: isAr ? 'نطاق التقرير' : 'Scope',
    vsBase: isAr ? 'مقارنة بسنة الأساس' : 'vs Base',
    segmentsLabel: isAr ? 'القطاعات' : 'Segments',
    allSegments: isAr ? 'كافة القطاعات' : 'ALL',
    grandTotal: isAr ? 'الإجمالي العام' : 'GRAND TOTAL',
    prior: isAr ? 'السابق' : 'Prior',
    current: isAr ? 'الحالي' : 'CURRENT',
    variance: isAr ? 'الفرق' : 'Variance',
    growth: isAr ? 'نسبة النمو' : 'Growth %',
    baseLabel: isAr ? 'الأساس' : 'Base',
    
    // Portal Action Buttons
    portalAnalytics: isAr ? 'الانتقال إلى تحليلات المبيعات' : 'Open Sales Analytics',
    portalPerformance: isAr ? 'فتح لوحة الأداء وقائمة الأسعار' : 'Open Performance & Price List',

    // Section 1: Executive Variance Breakdown by Segment
    sec1Title: isAr ? 'تحليل مبيعات الفترة لكل قطاع' : 'Executive Variance Breakdown by Segment',
    sec1Col1: isAr ? 'القطاع / التصنيف' : 'Segment / Category',
    sec1Empty: isAr ? 'لا توجد بيانات مبيعات مسجلة لنطاق هذه الفترة.' : 'No executive variance data recorded for this time scope.',

    // Section 2: Newly Late Clients Watchlist
    sec2Title: isAr ? 'قائمة مراقبة العملاء المتأخرين عن دورة الطلب المعتادة' : 'Newly Late Clients Watchlist',
    sec2Col1: isAr ? 'اسم العميل والفرع' : 'Client & Branch Name',
    sec2Col2: isAr ? 'آخر طلب' : 'Last Order',
    sec2Col3: isAr ? 'متوسط قيمة الطلب سابقاً' : 'Avg Order Qty',
    sec2Col4: isAr ? 'دورة الطلب المعتادة' : 'Typical Cycle',
    sec2Col5: isAr ? 'الحالة' : 'Status',
    sec2DaysOverdue: isAr ? 'يوم تأخير' : 'Days Overdue',
    sec2WatchlistBadge: isAr ? 'قائمة المراقبة' : 'Watchlist',
    sec2OverdueClients: isAr ? 'عميل متأخر' : 'Overdue Clients',
    sec2Empty: isAr ? 'لا يوجد عملاء متأخرون عن دورة الطلب في هذه الفترة.' : 'No late clients (2+ days overdue) detected in this segment filter.',

    // Section 3: Wins & Recovery
    sec3Title: isAr ? 'العملاء الجدد واستعادة العملاء المنقطعين' : 'Wins & Recovery (New & Re-engaged Clients)',
    sec3Col1: isAr ? 'اسم العميل' : 'Client Name',
    sec3Col2: isAr ? 'القطاع' : 'Segment',
    sec3Col3: isAr ? 'تصنيف العميل' : 'Classification',
    sec3Col4: isAr ? 'كمية الفترة' : 'Period Volume',
    sec3Col5: isAr ? 'مبيعات الفترة' : 'Period Revenue',
    sec3NewClient: isAr ? 'عميل جديد لأول مرة' : 'New Client',
    sec3Reengaged: isAr ? 'عميل مستعاد' : 'Re-engaged',
    sec3OverdueTag: isAr ? 'يوم انقطاع' : 'd Overdue',
    sec3Empty: isAr ? 'لا يوجد عملاء جدد أو مستعادون بعد انقطاع خلال هذه الفترة.' : 'No new or re-engaged (30+ days overdue) clients in this period.',

    // Section 4: Underperforming Products
    sec4Title: isAr ? 'المنتجات الأكثر تراجعاً في الأداء' : 'Underperforming Products vs Base Period',
    sec4Col1: isAr ? 'اسم الصنف / المنتج' : 'Product Name',
    sec4Col2: isAr ? 'الكمية السابقة' : 'Prior Vol',
    sec4Col3: isAr ? 'الكمية الحالية' : 'Current Vol',
    sec4Col4: isAr ? 'قيمة التراجع' : 'Volume Deficit',
    sec4Empty: isAr ? 'لم يتم رصد أي تراجع ملحوظ في مبيعات الأصناف خلال الفترة.' : 'All products in this scope are performing equal to or higher than base period!',

    // Section 5: Detailed Period Orders Register
    sec5Title: isAr ? 'سجل طلبات الفترة التفصيلي' : 'Detailed Period Orders Register',
    sec5Col1: isAr ? 'القطاع' : 'Segment',
    sec5Col2: isAr ? 'التاريخ' : 'Date',
    sec5Col3: isAr ? 'العميل والفرع' : 'Client & Branch',
    sec5Col4: isAr ? 'الأصناف والكميات المطلوبة' : 'Products & Quantities',
    sec5Col5: isAr ? 'إجمالي قيمة الطلب' : 'Order Value',
    sec5Col6: isAr ? 'طبيعة الطلب / الانتظام' : 'Consistency',
    sec5FirstOrder: isAr ? 'أول طلب للعميل' : 'First Order',
    sec5NormalVol: isAr ? 'كمية معتادة' : 'Normal Vol',
    sec5OnSchedule: isAr ? 'طلب منتظم' : 'On Schedule',
    sec5Delayed: isAr ? 'متأخر' : 'Delayed',
    sec5Early: isAr ? 'مبكر' : 'Early',
    sec5Empty: isAr ? 'لا توجد طلبات تم تسجيلها خلال هذه الفترة.' : 'No orders placed during this period.',

    // Section 6: Segment Period Orders Register
    sec6Title: isAr ? 'ملخص طلبات القطاعات' : 'Segment Period Orders Register',
    sec6Col1: isAr ? 'القطاع' : 'Segment Name',
    sec6Col2: isAr ? 'العملاء النشطين' : 'Active Clients',
    sec6Col3: isAr ? 'عدد الفواتير / الطلبات' : 'Total Orders',
    sec6Col4: isAr ? 'الأصناف والكميات المطلوبة' : 'Products & Quantities',
    sec6Col5: isAr ? 'إجمالي القيمة' : 'Total Revenue',
    sec6Empty: isAr ? 'لا توجد حركات بيع مسجلة للقطاعات خلال هذه الفترة.' : 'No segment activity recorded during this period.',

    // Section 7: Product Category Breakdown
    sec7Title: isAr ? 'تحليل أداء تصنيفات المنتجات' : 'Product Category Breakdown',
    sec7Col1: isAr ? 'تصنيف المنتج' : 'Product Category',
    sec7Empty: isAr ? 'لا توجد بيانات مسجلة لتصنيفات المنتجات خلال هذه الفترة.' : 'No product category breakdown data recorded for this time scope.',

    // Section 8: Macro Progress Cards
    sec8Title: isAr ? 'متابعة الأداء الكلي والتراكمي (MTD / QTD / YTD)' : 'Macro Progress & YTD Tracking',
    sec8RevHeader: isAr ? 'متابعة حركة الإيرادات (بالجنيه)' : 'REVENUE PROGRESS (EGP)',
    sec8VolHeader: isAr ? 'متابعة حركة الكميات (بالوحدة)' : 'VOLUME PROGRESS (UNITS)',
    sec8PeriodScope: isAr ? 'أداء الفترة' : 'Period Scope',
    sec8Mtd: isAr ? 'إجمالي الشهر حتى تاريخه (MTD)' : 'MTD Progress',
    sec8Qtd: isAr ? 'إجمالي الربع حتى تاريخه (QTD)' : 'QTD Progress',
    sec8Ytd: isAr ? 'الإجمالي التراكمي للسنة (YTD)' : 'YTD Macro Progress'
  };

  // Product Category Translation Dictionary
  var categoryDict = isAr ? {
    "WHITE T": "خل الطاووس أو برايفت ليبل",
    "other": "خل تفاح - ماء ورد - ماء زهر",
    "WHITE O": "خل أورجانو",
    "tahini": "طحينة",
    "Rice": "أرز"
  } : {
    "WHITE T": "White Vinegar (Tawoos / Pvt)",
    "other": "Apple / Rose / Blossom",
    "WHITE O": "Organo Vinegar",
    "tahini": "Tahini",
    "Rice": "Rice"
  };

  // Segment Filter Arrays (With Trimming Sanitization)
  var allowedSegs = (tpl.segments || ['ALL']).map(function(x) { return String(x).trim(); });
  var isAllSegs = allowedSegs.indexOf('ALL') !== -1;
  var isMatchSeg = function(s) { 
    if (isAllSegs) return true;
    var cleanS = String(s || '').trim();
    return allowedSegs.indexOf(cleanS) !== -1;
  };

  // Load JSON Datasets
  var rcData = (getRootCauseDataFromCache() || {}).data || {};
  var cadenceData = (getCadenceDataFromCache() || {}).data || [];
  var netSalesData = (getBranchNetSalesDataFromCache() || {}).data || [];
  var productData = (getProductDataFromCache() || {}).data || [];

  // ==========================================
  // DYNAMIC TIME SCOPE & BOUNDARY CALCULATION
  // ==========================================
  var now = new Date();
  var currYear = now.getFullYear();
  var currMonth = now.getMonth();
  var currDate = now.getDate();
  var baseYear = parseInt(tpl.baseYear) || (currYear - 1);

  var currStart, currEnd, baseStart, baseEnd;

  if (tpl.timeScope === 'Daily') {
    currStart = new Date(currYear, currMonth, currDate, 0, 0, 0);
    currEnd = new Date(currYear, currMonth, currDate, 23, 59, 59);
    baseStart = new Date(baseYear, currMonth, currDate, 0, 0, 0);
    baseEnd = new Date(baseYear, currMonth, currDate, 23, 59, 59);
  } else if (tpl.timeScope === 'Weekly') {
    currStart = new Date(currYear, currMonth, currDate - 6, 0, 0, 0);
    currEnd = new Date(currYear, currMonth, currDate, 23, 59, 59);
    baseStart = new Date(baseYear, currMonth, currDate - 6, 0, 0, 0);
    baseEnd = new Date(baseYear, currMonth, currDate, 23, 59, 59);
  } else if (tpl.timeScope === 'Monthly') {
    currStart = new Date(currYear, currMonth, 1, 0, 0, 0);
    currEnd = new Date(currYear, currMonth, currDate, 23, 59, 59);
    baseStart = new Date(baseYear, currMonth, 1, 0, 0, 0);
    baseEnd = new Date(baseYear, currMonth, currDate, 23, 59, 59);
  } else if (tpl.timeScope === 'Quarterly') {
    var q = Math.floor(currMonth / 3);
    var qDay = tpl.scheduleDay || 'Q_START';

    if (qDay === 'Q_END') {
      var qStartMonth = q * 3;
      var qEndMonth = qStartMonth + 2;
      var lastDayOfCurrQ = new Date(currYear, qEndMonth + 1, 0).getDate();
      var lastDayOfBaseQ = new Date(baseYear, qEndMonth + 1, 0).getDate();

      currStart = new Date(currYear, qStartMonth, 1, 0, 0, 0);
      currEnd = new Date(currYear, qEndMonth, lastDayOfCurrQ, 23, 59, 59);
      baseStart = new Date(baseYear, qStartMonth, 1, 0, 0, 0);
      baseEnd = new Date(baseYear, qEndMonth, lastDayOfBaseQ, 23, 59, 59);
    } else {
      currStart = new Date(currYear, q * 3, 1, 0, 0, 0);
      currEnd = new Date(currYear, currMonth, currDate, 23, 59, 59);
      baseStart = new Date(baseYear, q * 3, 1, 0, 0, 0);
      baseEnd = new Date(baseYear, currMonth, currDate, 23, 59, 59);
    }
  } else {
    currStart = new Date(currYear, 0, 1, 0, 0, 0);
    currEnd = new Date(currYear, currMonth, currDate, 23, 59, 59);
    baseStart = new Date(baseYear, 0, 1, 0, 0, 0);
    baseEnd = new Date(baseYear, currMonth, currDate, 23, 59, 59);
  }

  var fmtDate = function(d) {
    var dd = ('0' + d.getDate()).slice(-2);
    var mm = ('0' + (d.getMonth() + 1)).slice(-2);
    return dd + '/' + mm + '/' + d.getFullYear();
  };

  var currPeriodLabel = (tpl.timeScope === 'Daily') ? fmtDate(currStart) : (fmtDate(currStart) + ' - ' + fmtDate(currEnd));
  var basePeriodLabel = (tpl.timeScope === 'Daily') ? fmtDate(baseStart) : (fmtDate(baseStart) + ' - ' + fmtDate(baseEnd));

  // Dynamic Header Action Button
  var portalUrl = "";
  var portalBtnText = "";
  if (tpl.reportPortal === "PORTAL_ANALYTICS") {
    portalUrl = "https://script.google.com/macros/s/AKfycbzIjjcSJZeAdeW9HDz__thoM2V2ACW-LQMGlaEa_RdtzyEgkXV9BhMMopA_RNF794o8/exec";
    portalBtnText = i18n.portalAnalytics;
  } else if (tpl.reportPortal === "PORTAL_PERFORMANCE") {
    portalUrl = "https://script.google.com/macros/s/AKfycbyTjgPM-Rr5KeJCovani-eghHWwgTQsgvfDQ6Y7CIEnItXJlmy_MtKaU21b1gLQP91cFQ/exec";
    portalBtnText = i18n.portalPerformance;
  }

  var btnHtml = "";
  if (portalUrl) {
    btnHtml = `
      <a href="${portalUrl}" target="_blank" style="display:inline-block; background-color:#ffffff; color:#003366; padding:10px 18px; text-decoration:none; border-radius:8px; font-weight:800; font-size:0.8rem; box-shadow:0 2px 4px rgba(0,0,0,0.15);">
        ${portalBtnText}
      </a>
    `;
  }

  var displaySegs = isAllSegs ? i18n.allSegments : allowedSegs.join(', ');

  var html = `
  <!DOCTYPE html>
  <html dir="${dir}">
  <head>
    <meta charset="utf-8">
    <style>
      body { font-family: 'Segoe UI', Tahoma, Arial, sans-serif; background-color: #f8fafc; color: #1e293b; padding: 20px; direction: ${dir}; text-align: ${align}; }
      .card { background: #ffffff; border-radius: 12px; padding: 20px; margin-bottom: 20px; border: 1px solid #e2e8f0; }
      .header { background: #003366; color: white; padding: 20px; border-radius: 12px 12px 0 0; }
      table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 0.85rem; }
      th { background-color: #f1f5f9; color: #475569; text-align: center; padding: 8px; border-bottom: 2px solid #cbd5e1; }
      td { padding: 8px; border-bottom: 1px solid #e2e8f0; text-align: center; }
      .badge { display: inline-block; padding: 3px 8px; border-radius: 12px; font-size: 0.75rem; font-weight: bold; }
      .badge-success { background: #dcfce7; color: #15803d; }
      .badge-danger { background: #fee2e2; color: #b91c1c; }
      .badge-warning { background: #fef3c7; color: #b45309; }
    </style>
  </head>
  <body dir="${dir}">
    <div class="card" style="padding:0; overflow:hidden;">
      <div class="header">
        <table style="width:100%; border-collapse:collapse; margin:0; font-family:'Segoe UI', Arial, sans-serif;">
          <tr>
            <td style="vertical-align:middle; border:none; padding:0; text-align:${align};">
              <h2 style="margin:0; color:#ffffff; font-size:1.3rem;">${tpl.templateName}</h2>
              <p style="margin:5px 0 0 0; opacity:0.9; font-size:0.85rem; font-weight: bold; color:#ffffff;">
                ${i18n.scopeLabel}: ${tpl.timeScope} (${currPeriodLabel}) ${i18n.vsBase} (${basePeriodLabel})
              </p>
              <p style="margin:3px 0 0 0; opacity:0.75; font-size:0.75rem; color:#ffffff;">
                ${i18n.segmentsLabel}: ${displaySegs}
              </p>
            </td>
            <td style="vertical-align:middle; border:none; padding:0; text-align:${isAr ? 'left' : 'right'}; min-width:180px;">
              ${btnHtml}
            </td>
          </tr>
        </table>
      </div>
      <div style="padding:20px;">
  `;

  // ==========================================
  // DYNAMIC REORDERABLE SECTIONS ENGINE
  // ==========================================
  var currStartMs = currStart.getTime();
  var currEndMs = currEnd.getTime();
  var baseStartMs = baseStart.getTime();
  var baseEndMs = baseEnd.getTime();

  var sectionCounter = 1;
  var activeSectionsList = Array.isArray(tpl.sections) ? tpl.sections.map(Number) : [1, 2, 3, 4, 5, 6, 7, 8];

  activeSectionsList.forEach(function(secId) {

    // --- SECTION 1: EXECUTIVE VARIANCE BREAKDOWN BY SEGMENT ---
    if (secId === 1) {
      html += `
        <div style="background:#ffffff; border:1px solid #cbd5e1; border-${isAr ? 'right' : 'left'}:6px solid #003366; border-radius:12px; padding:20px; margin-bottom:30px; box-shadow:0 4px 6px rgba(0,0,0,0.03);">
          <div style="background-color:#003366; color:#ffffff; padding:10px 16px; border-radius:8px; font-weight:bold; font-size:1.05rem; margin-bottom:15px;">
            ${sectionCounter}. ${i18n.sec1Title}
          </div>
      `;

      var segMap = {};
      var grandTotal = { currV: 0, prevV: 0, currQ: 0, prevQ: 0 };

      Object.keys(rcData).forEach(function(dStr) {
        var dMs = new Date(dStr).getTime();
        var isCurr = (dMs >= currStartMs && dMs <= currEndMs);
        var isPrev = (dMs >= baseStartMs && dMs <= baseEndMs);

        if (isCurr || isPrev) {
          var segs = rcData[dStr];
          Object.keys(segs).forEach(function(s) {
            if (!isMatchSeg(s)) return;
            if (!segMap[s]) segMap[s] = { currV: 0, prevV: 0, currQ: 0, prevQ: 0 };
            Object.keys(segs[s]).forEach(function(m) {
              Object.keys(segs[s][m]).forEach(function(b) {
                Object.keys(segs[s][m][b]).forEach(function(c) {
                  Object.keys(segs[s][m][b][c]).forEach(function(t) {
                    Object.keys(segs[s][m][b][c][t]).forEach(function(p) {
                      var item = segs[s][m][b][c][t][p];
                      var q = item.q || 0;
                      var v = item.v || 0;

                      if (isCurr) {
                        segMap[s].currV += v;
                        segMap[s].currQ += q;
                        grandTotal.currV += v;
                        grandTotal.currQ += q;
                      }
                      if (isPrev) {
                        segMap[s].prevV += v;
                        segMap[s].prevQ += q;
                        grandTotal.prevV += v;
                        grandTotal.prevQ += q;
                      }
                    });
                  });
                });
              });
            });
          });
        }
      });

      // Sort segments by Current Revenue descending
      var segKeys = Object.keys(segMap).sort(function(a, b) {
        return (segMap[b].currV || 0) - (segMap[a].currV || 0);
      });

      var renderCompositeCell = function(prevV, currV, prevQ, currQ, colType, isCurrentCol) {
        var diffV = currV - prevV;
        var diffQ = currQ - prevQ;

        var pctV = prevV > 0 ? ((diffV / prevV) * 100) : (currV > 0 ? 100 : 0);
        var pctQ = prevQ > 0 ? ((diffQ / prevQ) * 100) : (currQ > 0 ? 100 : 0);

        var colorV = diffV >= 0 ? "#15803d" : "#b91c1c";
        var colorQ = diffQ >= 0 ? "#15803d" : "#b91c1c";
        var bgQ = diffQ >= 0 ? "#dcfce7" : "#fee2e2";

        var signV = diffV >= 0 ? "+" : "-";
        var signQ = diffQ >= 0 ? "+" : "-";

        var revHtml = "";
        var volHtml = "";

        if (colType === "PRIOR") {
          revHtml = `<strong style="color:#475569;">${formatKM(prevV, revUnit)}</strong>`;
          volHtml = `<span style="display:inline-block; font-size:0.72rem; background:#e2e8f0; color:#475569; padding:2px 6px; border-radius:10px; margin-${isAr?'right':'left'}:5px; font-weight:600; white-space:nowrap;">${formatKM(prevQ, qtyUnit)}</span>`;

        } else if (colType === "CURRENT") {
          revHtml = `<strong style="color:#003366; font-size:0.92rem;">${formatKM(currV, revUnit)}</strong>`;
          volHtml = `<span style="display:inline-block; font-size:0.72rem; background:#cbd5e1; color:#0f172a; padding:2px 6px; border-radius:10px; margin-${isAr?'right':'left'}:5px; font-weight:bold; white-space:nowrap;">${formatKM(currQ, qtyUnit)}</span>`;

        } else if (colType === "VARIANCE") {
          revHtml = `<strong style="color:${colorV};">${signV}${formatKM(Math.abs(diffV), revUnit)}</strong>`;
          volHtml = `<span style="display:inline-block; font-size:0.72rem; background:${bgQ}; color:${colorQ}; padding:2px 6px; border-radius:10px; margin-${isAr?'right':'left'}:5px; font-weight:bold; white-space:nowrap;">${signQ}${formatKM(Math.abs(diffQ), qtyUnit)}</span>`;

        } else if (colType === "GROWTH") {
          revHtml = `<strong style="color:${colorV};">${signV}${Math.abs(pctV).toFixed(1)}%</strong>`;
          volHtml = `<span style="display:inline-block; font-size:0.72rem; background:${bgQ}; color:${colorQ}; padding:2px 6px; border-radius:10px; margin-${isAr?'right':'left'}:5px; font-weight:bold; white-space:nowrap;">${signQ}${Math.abs(pctQ).toFixed(1)}%</span>`;
        }

        var cellBg = isCurrentCol ? "background-color:#eff6ff;" : "";
        return `<td style="padding:9px 10px; text-align:center; border-right:1px solid #cbd5e1; ${cellBg}">${revHtml}${volHtml}</td>`;
      };

      html += `
        <div style="margin-top:10px; margin-bottom:10px; border:1px solid #cbd5e1; border-radius:8px; overflow:hidden;">
          <table style="width:100%; border-collapse:collapse; font-size:0.85rem; font-family:'Segoe UI', Arial, sans-serif;">
            <thead>
              <tr style="background-color:#f8fafc; color:#475569;">
                <th style="padding:10px; text-align:center; border-bottom:2px solid #003366; border-right:1px solid #cbd5e1; min-width:140px;">${i18n.sec1Col1}</th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #003366; border-right:1px solid #cbd5e1; color:#64748b; background-color:#f1f5f9;">${i18n.prior} (${basePeriodLabel})</th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #003366; border-right:1px solid #cbd5e1; color:#ffffff; background-color:#003366; font-size:0.9rem;">${i18n.current} (${currPeriodLabel})</th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #003366; border-right:1px solid #cbd5e1; background-color:#f1f5f9;">${i18n.variance}</th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #003366; background-color:#f1f5f9;">${i18n.growth}</th>
              </tr>
            </thead>
            <tbody>
      `;

      if (segKeys.length === 0) {
        html += `<tr><td colspan="5" style="text-align:center; color:#64748b; padding:15px;">${i18n.sec1Empty}</td></tr>`;
      } else {
        var renderRow = function(name, prevV, currV, prevQ, currQ, isTotal) {
          var bgRow = isTotal ? "background-color:#e2e8f0; font-weight:bold; border-top:2px solid #003366;" : "border-top:1px solid #e2e8f0;";
          var labelText = isTotal 
            ? `<strong style="color:#003366; font-size:0.95rem;">${i18n.grandTotal}</strong>` 
            : `<strong style="color:#1e293b;">${name}</strong>`;

          return `
            <tr style="${bgRow}">
              <td style="padding:9px 10px; text-align:center; border-right:1px solid #cbd5e1;">${labelText}</td>
              ${renderCompositeCell(prevV, currV, prevQ, currQ, "PRIOR", false)}
              ${renderCompositeCell(prevV, currV, prevQ, currQ, "CURRENT", true)}
              ${renderCompositeCell(prevV, currV, prevQ, currQ, "VARIANCE", false)}
              ${renderCompositeCell(prevV, currV, prevQ, currQ, "GROWTH", false)}
            </tr>`;
        };

        html += renderRow(i18n.grandTotal, grandTotal.prevV, grandTotal.currV, grandTotal.prevQ, grandTotal.currQ, true);

        segKeys.forEach(function(sName) {
          var sData = segMap[sName];
          html += renderRow(sName, sData.prevV, sData.currV, sData.prevQ, sData.currQ, false);
        });
      }

      html += `</tbody></table></div></div>`;
      sectionCounter++;
    }

    // --- SECTION 2: NEWLY LATE CLIENTS WATCHLIST ---
    else if (secId === 2) {
      html += `
        <div style="background:#ffffff; border:1px solid #cbd5e1; border-${isAr ? 'right' : 'left'}:6px solid #2563eb; border-radius:12px; padding:20px; margin-bottom:30px; box-shadow:0 4px 6px rgba(0,0,0,0.03);">
          <div style="background-color:#1e40af; color:#ffffff; padding:10px 16px; border-radius:8px; font-weight:bold; font-size:1.05rem; margin-bottom:15px;">
            ${sectionCounter}. ${i18n.sec2Title}
          </div>
      `;

      var alerts = processEmailClientCadence(cadenceData).filter(function(a) { return isMatchSeg(a.segment); });

      if (alerts.length === 0) {
        html += `<div style="padding:15px; border:1px solid #cbd5e1; border-radius:8px; text-align:center; color:#64748b;">${i18n.sec2Empty}</div>`;
      } else {
        var groupedAlerts = {};
        alerts.forEach(function(a) {
          var seg = a.segment || 'Unassigned';
          if (!groupedAlerts[seg]) groupedAlerts[seg] = [];
          groupedAlerts[seg].push(a);
        });

        Object.keys(groupedAlerts).forEach(function(segName) {
          var segAlerts = groupedAlerts[segName];
          segAlerts.sort(function(a, b) { return b.avgOrderQty - a.avgOrderQty; });

          html += `
            <div style="margin-top:10px; margin-bottom:20px; border:1px solid #cbd5e1; border-radius:8px; overflow:hidden;">
              <div style="background-color:#1e40af; color:#ffffff; padding:8px 14px; font-weight:bold; font-size:0.85rem; text-align:center;">
                ${segName} - ${i18n.sec2WatchlistBadge} (${segAlerts.length} ${i18n.sec2OverdueClients})
              </div>
              <table style="width:100%; border-collapse:collapse; font-size:0.85rem; font-family:'Segoe UI', Arial, sans-serif;">
                <thead>
                  <tr style="background-color:#f8fafc; color:#475569;">
                    <th style="padding:10px; text-align:center; border-bottom:2px solid #1e40af; border-right:1px solid #cbd5e1;">${i18n.sec2Col1}</th>
                    <th style="padding:10px; text-align:center; border-bottom:2px solid #1e40af; border-right:1px solid #cbd5e1;">${i18n.sec2Col2}</th>
                    <th style="padding:10px; text-align:center; border-bottom:2px solid #1e40af; border-right:1px solid #cbd5e1;">${i18n.sec2Col3}</th>
                    <th style="padding:10px; text-align:center; border-bottom:2px solid #1e40af; border-right:1px solid #cbd5e1;">${i18n.sec2Col4}</th>
                    <th style="padding:10px; text-align:center; border-bottom:2px solid #1e40af;">${i18n.sec2Col5}</th>
                  </tr>
                </thead>
                <tbody>
          `;

          segAlerts.forEach(function(a) {
            var formattedLastOrderDate = fmtDate(new Date(a.lastOrderDate));
            html += `
              <tr style="border-top:1px solid #e2e8f0;">
                <td style="padding:9px 10px; text-align:center; border-right:1px solid #cbd5e1;"><strong>${a.client}</strong></td>
                <td style="padding:9px 10px; text-align:center; border-right:1px solid #cbd5e1;">
                  <div style="font-weight:bold; color:#1e293b;">${formatKM(a.lastOrderQty, qtyUnit)}</div>
                  <div style="font-size:0.75rem; color:#64748b; margin-top:2px;">${formattedLastOrderDate}</div>
                </td>
                <td style="padding:9px 10px; text-align:center; border-right:1px solid #cbd5e1; font-weight:bold; color:#003366;">${formatKM(a.avgOrderQty, qtyUnit)}</td>
                <td style="padding:9px 10px; text-align:center; border-right:1px solid #cbd5e1;">~${a.expectedInterval} ${isAr ? 'يوم' : 'Days'}</td>
                <td style="padding:9px 10px; text-align:center;"><span class="badge badge-danger">${a.daysLate} ${i18n.sec2DaysOverdue}</span></td>
              </tr>`;
          });
          html += `</tbody></table></div>`;
        });
      }

      html += `</div>`;
      sectionCounter++;
    }

    // --- SECTION 3: WINS & RECOVERY ---
    else if (secId === 3) {
      html += `
        <div style="background:#ffffff; border:1px solid #cbd5e1; border-${isAr ? 'right' : 'left'}:6px solid #16a34a; border-radius:12px; padding:20px; margin-bottom:30px; box-shadow:0 4px 6px rgba(0,0,0,0.03);">
          <div style="background-color:#15803d; color:#ffffff; padding:10px 16px; border-radius:8px; font-weight:bold; font-size:1.05rem; margin-bottom:15px;">
            ${sectionCounter}. ${i18n.sec3Title}
          </div>
          <div style="border:1px solid #cbd5e1; border-radius:8px; overflow:hidden;">
          <table style="width:100%; border-collapse:collapse; font-size:0.85rem; font-family:'Segoe UI', Arial, sans-serif;">
            <thead>
              <tr style="background-color:#f8fafc; color:#475569;">
                <th style="padding:10px; text-align:center; border-bottom:2px solid #15803d; border-right:1px solid #cbd5e1;">${i18n.sec3Col1}</th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #15803d; border-right:1px solid #cbd5e1;">${i18n.sec3Col2}</th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #15803d; border-right:1px solid #cbd5e1;">${i18n.sec3Col3}</th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #15803d; border-right:1px solid #cbd5e1;">${i18n.sec3Col4}</th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #15803d;">${i18n.sec3Col5}</th>
              </tr>
            </thead>
            <tbody>
      `;

      var msPerDay = 1000 * 60 * 60 * 24;
      var clientSegmentLookup = {};
      cadenceData.forEach(function(c) {
        if (c.client && c.segment && c.segment !== 'Unassigned') {
          clientSegmentLookup[String(c.client).trim()] = String(c.segment).trim();
        }
      });

      var clientHistoryMap = {};
      netSalesData.forEach(function(s) {
        var client = String(s.client || s.Client || '').trim();
        if (!client) return;

        var dMs = parseDateMs(s.dateMs || s.date || s.Date);
        if (!dMs) return;

        var d = new Date(dMs);
        d.setHours(0, 0, 0, 0);
        var dateKey = d.getTime();

        var rawSeg = String(s.segment || s.Segment || '').trim();
        var seg = (rawSeg && rawSeg !== 'Unassigned') ? rawSeg : (clientSegmentLookup[client] || 'Unassigned');

        if (!clientHistoryMap[client]) clientHistoryMap[client] = { ordersByDate: {} };
        if (!clientHistoryMap[client].ordersByDate[dateKey]) {
          clientHistoryMap[client].ordersByDate[dateKey] = { dateMs: dateKey, qty: 0, rev: 0, segment: seg };
        }

        clientHistoryMap[client].ordersByDate[dateKey].qty += (Number(s.qty || s.Qty || s.quantity) || 0);
        clientHistoryMap[client].ordersByDate[dateKey].rev += (Number(s.netRev || s.rev || s.revenue) || 0);
        if (seg !== 'Unassigned') clientHistoryMap[client].ordersByDate[dateKey].segment = seg;
      });

      var winRecords = [];

      Object.keys(clientHistoryMap).forEach(function(client) {
        var ordersObj = clientHistoryMap[client].ordersByDate;
        var dateKeys = Object.keys(ordersObj);
        var sortedDates = dateKeys.map(Number).sort(function(a, b) { return a - b; });
        if (sortedDates.length === 0) return;

        var windowDates = sortedDates.filter(function(d) { return d >= currStartMs && d <= currEndMs; });
        if (windowDates.length === 0) return;

        var currentOrderSeg = ordersObj[windowDates[0]].segment;
        if (currentOrderSeg === 'Unassigned' && clientSegmentLookup[client]) {
          currentOrderSeg = clientSegmentLookup[client];
        }

        if (!isMatchSeg(currentOrderSeg)) return;

        var windowQty = 0; var windowRev = 0;
        windowDates.forEach(function(d) {
          windowQty += ordersObj[d].qty;
          windowRev += ordersObj[d].rev;
        });

        var firstEverDate = sortedDates[0];
        var currentFirstOrderDate = windowDates[0];

        if (firstEverDate >= currStartMs && firstEverDate <= currEndMs) {
          winRecords.push({
            client: client, segment: currentOrderSeg, badgeText: i18n.sec3NewClient,
            badgeStyle: 'background-color:#dcfce7; color:#15803d;', qty: windowQty, rev: windowRev
          });
        } else {
          var priorDates = sortedDates.filter(function(d) { return d < currentFirstOrderDate; });

          if (priorDates.length > 0) {
            var lastPriorDate = priorDates[priorDates.length - 1];
            var gapDays = (currentFirstOrderDate - lastPriorDate) / msPerDay;

            var avgCycleDays = 30;
            var recentPriorDates = priorDates.slice(-4);

            if (recentPriorDates.length >= 2) {
              var priorFirst = recentPriorDates[0];
              var elapsed = (lastPriorDate - priorFirst) / msPerDay;
              var priorQtySum = 0;
              for (var k = 0; k < recentPriorDates.length - 1; k++) {
                priorQtySum += (ordersObj[recentPriorDates[k]].qty || 0);
              }

              if (elapsed > 0 && priorQtySum > 0) {
                var dailyRate = priorQtySum / elapsed;
                var latestPriorQty = ordersObj[lastPriorDate].qty || 0;
                if (dailyRate > 0 && latestPriorQty > 0) {
                  avgCycleDays = Math.min(60, latestPriorQty / dailyRate);
                }
              }
            }

            var daysOverdueBeforePurchase = gapDays - avgCycleDays;

            if (daysOverdueBeforePurchase >= 30 || gapDays >= 60) {
              var overdueText = Math.floor(Math.max(daysOverdueBeforePurchase, gapDays - 30));
              var badgeStr = isAr 
                ? `${i18n.sec3Reengaged} (+${overdueText} ${i18n.sec3OverdueTag})`
                : `${i18n.sec3Reengaged} (+${overdueText}${i18n.sec3OverdueTag})`;

              winRecords.push({
                client: client, segment: currentOrderSeg, badgeText: badgeStr,
                badgeStyle: 'background-color:#dbeafe; color:#1e40af;', qty: windowQty, rev: windowRev
              });
            }
          }
        }
      });

      winRecords.sort(function(a, b) { return b.rev - a.rev; });

      if (winRecords.length === 0) {
        html += `<tr><td colspan="5" style="text-align:center; color:#64748b; padding:15px;">${i18n.sec3Empty}</td></tr>`;
      } else {
        winRecords.slice(0, 10).forEach(function(w) {
          html += `
            <tr style="border-top:1px solid #e2e8f0;">
              <td style="padding:9px 10px; text-align:center; border-right:1px solid #cbd5e1;"><strong>${w.client}</strong></td>
              <td style="padding:9px 10px; text-align:center; border-right:1px solid #cbd5e1;"><span class="badge badge-warning">${w.segment}</span></td>
              <td style="padding:9px 10px; text-align:center; border-right:1px solid #cbd5e1;"><span class="badge" style="${w.badgeStyle}">${w.badgeText}</span></td>
              <td style="padding:9px 10px; text-align:center; border-right:1px solid #cbd5e1; font-weight:bold; color:#1e293b;">${formatKM(w.qty, qtyUnit)}</td>
              <td style="padding:9px 10px; text-align:center; font-weight:bold; color:#15803d;">${formatKM(w.rev, revUnit)}</td>
            </tr>`;
        });
      }

      html += `</tbody></table></div></div>`;
      sectionCounter++;
    }

    // --- SECTION 4: UNDERPERFORMING PRODUCTS ---
    else if (secId === 4) {
      html += `
        <div style="background:#ffffff; border:1px solid #cbd5e1; border-${isAr ? 'right' : 'left'}:6px solid #dc2626; border-radius:12px; padding:20px; margin-bottom:30px; box-shadow:0 4px 6px rgba(0,0,0,0.03);">
          <div style="background-color:#b91c1c; color:#ffffff; padding:10px 16px; border-radius:8px; font-weight:bold; font-size:1.05rem; margin-bottom:15px;">
            ${sectionCounter}. ${i18n.sec4Title}
          </div>
          <table style="width:100%; border-collapse:collapse; font-size:0.85rem; font-family:'Segoe UI', Arial, sans-serif;">
            <thead>
              <tr style="background-color:#f8fafc; color:#475569;">
                <th style="padding:10px; text-align:center; border-bottom:2px solid #b91c1c; border-right:1px solid #cbd5e1;">${i18n.sec4Col1}</th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #b91c1c; border-right:1px solid #cbd5e1;">${i18n.sec4Col2}</th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #b91c1c; border-right:1px solid #cbd5e1;">${i18n.sec4Col3}</th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #b91c1c;">${i18n.sec4Col4}</th>
              </tr>
            </thead>
            <tbody>
      `;

      var prodMap = {};
      productData.forEach(function(p) {
        var pMs = new Date(p.date).getTime();
        var isCurr = (pMs >= currStartMs && pMs <= currEndMs);
        var isPrev = (pMs >= baseStartMs && pMs <= baseEndMs);

        if (isCurr || isPrev) {
          var hasSegVol = isAllSegs;
          if (!hasSegVol && p.segments) {
            allowedSegs.forEach(function(sg) { if (p.segments[sg]) hasSegVol = true; });
          }
          if (hasSegVol) {
            if (!prodMap[p.item]) prodMap[p.item] = { name: p.item, curr: 0, prev: 0 };
            if (isCurr) prodMap[p.item].curr += p.qty;
            if (isPrev) prodMap[p.item].prev += p.qty;
          }
        }
      });

      var underperforming = Object.values(prodMap).filter(function(p){ return p.curr < p.prev; }).sort(function(a,b){ return (a.curr - a.prev) - (b.curr - b.prev); }).slice(0, 5);

      if (underperforming.length === 0) {
        html += `<tr><td colspan="4" style="text-align:center; color:#64748b; padding:15px;">${i18n.sec4Empty}</td></tr>`;
      } else {
        underperforming.forEach(function(up) {
          var deficit = up.curr - up.prev;
          html += `
            <tr style="border-top:1px solid #e2e8f0;">
              <td style="padding:9px 10px; text-align:center; border-right:1px solid #cbd5e1;"><strong>${up.name}</strong></td>
              <td style="padding:9px 10px; text-align:center; border-right:1px solid #cbd5e1;">${formatKM(up.prev, qtyUnit)}</td>
              <td style="padding:9px 10px; text-align:center; border-right:1px solid #cbd5e1;">${formatKM(up.curr, qtyUnit)}</td>
              <td style="padding:9px 10px; text-align:center; color:#b91c1c; font-weight:bold;">${formatKM(deficit, qtyUnit)}</td>
            </tr>`;
        });
      }

      html += `</tbody></table></div>`;
      sectionCounter++;
    }

    // --- SECTION 5: DETAILED PERIOD ORDERS REGISTER ---
    else if (secId === 5) {
      html += `
        <div style="background:#ffffff; border:1px solid #cbd5e1; border-${isAr ? 'right' : 'left'}:6px solid #8b5cf6; border-radius:12px; padding:20px; margin-bottom:30px; box-shadow:0 4px 6px rgba(0,0,0,0.03);">
          <div style="background-color:#6d28d9; color:#ffffff; padding:10px 16px; border-radius:8px; font-weight:bold; font-size:1.05rem; margin-bottom:15px;">
            ${sectionCounter}. ${i18n.sec5Title}
          </div>
      `;

      if (!netSalesData || netSalesData.length === 0) {
        if (typeof generateBranchNetSalesCache === 'function') {
          netSalesData = generateBranchNetSalesCache(true) || [];
        }
      }

      var msPerDay = 1000 * 60 * 60 * 24;
      var clientSegmentLookup = {};
      if (Array.isArray(cadenceData)) {
        cadenceData.forEach(function(c) {
          if (c.client && c.segment && c.segment !== 'Unassigned') {
            clientSegmentLookup[String(c.client).trim()] = String(c.segment).trim();
          }
        });
      }

      var allClientHistory = {};
      netSalesData.forEach(function(s) {
        var client = String(s.client || s.Client || '').trim();
        if (!client) return;
        var dMs = parseDateMs(s.dateMs || s.date || s.Date);
        if (!dMs) return;

        var d = new Date(dMs);
        d.setHours(0, 0, 0, 0);
        var dateKey = d.getTime();

        var rawSeg = String(s.segment || s.Segment || '').trim();
        var seg = (rawSeg && rawSeg !== 'Unassigned') ? rawSeg : (clientSegmentLookup[client] || 'Unassigned');

        if (!allClientHistory[client]) allClientHistory[client] = {};
        if (!allClientHistory[client][dateKey]) {
          allClientHistory[client][dateKey] = {
            date: dateKey, totalQty: 0, totalRev: 0,
            segment: seg, items: {}
          };
        }

        var item = String(s.item || s.product || s.Product || 'Item').trim();
        var q = Number(s.qty || s.Qty || 0);
        var r = Number(s.netRev || s.rev || 0);

        allClientHistory[client][dateKey].totalQty += q;
        allClientHistory[client][dateKey].totalRev += r;
        allClientHistory[client][dateKey].items[item] = (allClientHistory[client][dateKey].items[item] || 0) + q;
      });

      var segmentOrdersMap = {};
      Object.keys(allClientHistory).forEach(function(client) {
        var dateKeys = Object.keys(allClientHistory[client]).map(Number).sort(function(a, b) { return a - b; });

        dateKeys.forEach(function(dKey) {
          if (dKey >= currStartMs && dKey <= currEndMs) {
            var orderObj = allClientHistory[client][dKey];
            var seg = orderObj.segment;

            if (!isMatchSeg(seg)) return;

            var priorOrders = dateKeys.filter(function(k) { return k < dKey; }).map(function(k) {
              return allClientHistory[client][k];
            });

            var consistency = calculateOrderConsistency(dKey, orderObj.totalQty, priorOrders, msPerDay);

            // Multilingual consistency badges
            var badgeTextTranslated = consistency.text;
            if (isAr) {
              badgeTextTranslated = badgeTextTranslated
                .replace("First Order", i18n.sec5FirstOrder)
                .replace("Normal Vol", i18n.sec5NormalVol)
                .replace("On Schedule", i18n.sec5OnSchedule)
                .replace("Delayed", i18n.sec5Delayed)
                .replace("Early", i18n.sec5Early)
                .replace("Vol", "كمية")
                .replace("d", " يوم");
            }

            var itemListStr = Object.keys(orderObj.items).map(function(pName) {
              var qty = orderObj.items[pName];
              return '<div style="margin-bottom:3px;"><strong style="color:#003366;">' + qty + 'x</strong> ' + pName + '</div>';
            }).join("");

            if (!segmentOrdersMap[seg]) segmentOrdersMap[seg] = [];

            segmentOrdersMap[seg].push({
              dateMs: dKey, dateStr: fmtDate(new Date(dKey)), client: client,
              productsStr: itemListStr, totalQty: orderObj.totalQty, totalRev: orderObj.totalRev,
              consistencyText: badgeTextTranslated, consistencyStyle: consistency.style
            });
          }
        });
      });

      var activeSegKeys = Object.keys(segmentOrdersMap);
      activeSegKeys.sort(function(a, b) {
        var revA = segmentOrdersMap[a].reduce(function(sum, o) { return sum + o.totalRev; }, 0);
        var revB = segmentOrdersMap[b].reduce(function(sum, o) { return sum + o.totalRev; }, 0);
        return revB - revA;
      });

      if (activeSegKeys.length === 0) {
        html += `<div style="padding:15px; border:1px solid #cbd5e1; border-radius:8px; text-align:center; color:#64748b;">${i18n.sec5Empty}</div>`;
      } else {
        html += `
          <div style="border:1px solid #cbd5e1; border-radius:8px; overflow:hidden;">
            <table style="width:100%; border-collapse:collapse; font-size:0.85rem; font-family:'Segoe UI', Arial, sans-serif;">
              <thead>
                <tr style="background-color:#f8fafc; color:#475569;">
                  <th style="padding:10px; text-align:center; border-bottom:2px solid #6d28d9; border-right:1px solid #cbd5e1; width:80px;">${i18n.sec5Col1}</th>
                  <th style="padding:10px; text-align:center; border-bottom:2px solid #6d28d9; border-right:1px solid #cbd5e1; width:85px;">${i18n.sec5Col2}</th>
                  <th style="padding:10px; text-align:center; border-bottom:2px solid #6d28d9; border-right:1px solid #cbd5e1; min-width:130px;">${i18n.sec5Col3}</th>
                  <th style="padding:10px; text-align:center; border-bottom:2px solid #6d28d9; border-right:1px solid #cbd5e1;">${i18n.sec5Col4}</th>
                  <th style="padding:10px; text-align:center; border-bottom:2px solid #6d28d9; border-right:1px solid #cbd5e1; width:100px;">${i18n.sec5Col5}</th>
                  <th style="padding:10px; text-align:center; border-bottom:2px solid #6d28d9; width:140px;">${i18n.sec5Col6}</th>
                </tr>
              </thead>
              <tbody>
        `;

        activeSegKeys.forEach(function(segName) {
          var ordersList = segmentOrdersMap[segName];
          ordersList.sort(function(a, b) { return b.dateMs - a.dateMs; });

          var rowCount = ordersList.length;

          ordersList.forEach(function(o, idx) {
            var isFirstRow = (idx === 0);
            var isLastRowOfSeg = (idx === rowCount - 1);
            var borderBottomStyle = isLastRowOfSeg ? "border-bottom:2px solid #6d28d9;" : "border-bottom:1px solid #e2e8f0;";

            html += `<tr style="${borderBottomStyle}">`;

            if (isFirstRow) {
              var chipStyle = "display:inline-block; padding:5px 10px; white-space:nowrap;";
              html += `
                <td rowspan="${rowCount}" style="padding:8px; text-align:center; vertical-align:middle; background-color:#f8fafc; border-right:1px solid #cbd5e1; ${borderBottomStyle}">
                  <span class="badge badge-warning" style="${chipStyle}">${segName}</span>
                </td>
              `;
            }

            html += `
                <td style="padding:9px 10px; text-align:center; border-right:1px solid #cbd5e1; color:#64748b; font-size:0.8rem;">${o.dateStr}</td>
                <td style="padding:9px 10px; text-align:center; border-right:1px solid #cbd5e1;"><strong>${o.client}</strong></td>
                <td style="padding:9px 10px; text-align:center; border-right:1px solid #cbd5e1; color:#334155; font-size:0.8rem;">${o.productsStr}</td>
                <td style="padding:9px 10px; text-align:center; border-right:1px solid #cbd5e1; font-weight:bold; color:#15803d;">${formatKM(o.totalRev, revUnit)}</td>
                <td style="padding:9px 10px; text-align:center;"><span class="badge" style="${o.consistencyStyle}">${o.consistencyText}</span></td>
              </tr>
            `;
          });
        });

        html += `</tbody></table></div>`;
      }

      html += `</div>`;
      sectionCounter++;
    }

    // --- SECTION 6: SEGMENT PERIOD ORDERS REGISTER ---
    else if (secId === 6) {
      html += `
        <div style="background:#ffffff; border:1px solid #cbd5e1; border-${isAr ? 'right' : 'left'}:6px solid #0f766e; border-radius:12px; padding:20px; margin-bottom:30px; box-shadow:0 4px 6px rgba(0,0,0,0.03);">
          <div style="background-color:#0f766e; color:#ffffff; padding:10px 16px; border-radius:8px; font-weight:bold; font-size:1.05rem; margin-bottom:15px;">
            ${sectionCounter}. ${i18n.sec6Title}
          </div>
      `;

      var segSummaryMap = {};

      netSalesData.forEach(function(s) {
        var client = String(s.client || s.Client || '').trim();
        if (!client) return;

        var dMs = parseDateMs(s.dateMs || s.date || s.Date);
        if (!dMs) return;

        var d = new Date(dMs);
        d.setHours(0, 0, 0, 0);
        var dateKey = d.getTime();

        if (dateKey >= currStartMs && dateKey <= currEndMs) {
          var rawSeg = String(s.segment || s.Segment || 'Unassigned').trim();
          var seg = (rawSeg && rawSeg !== 'Unassigned') ? rawSeg : (clientSegmentLookup[client] || 'Unassigned');

          if (!isMatchSeg(seg)) return;

          if (!segSummaryMap[seg]) {
            segSummaryMap[seg] = {
              clientsMap: {},
              ordersMap: {},
              items: {},
              totalRev: 0
            };
          }

          var item = String(s.item || s.product || s.Product || 'Item').trim();
          var q = Number(s.qty || s.Qty || 0);
          var r = Number(s.netRev || s.rev || 0);

          segSummaryMap[seg].clientsMap[client] = true;
          segSummaryMap[seg].ordersMap[client + "_" + dateKey] = true;
          segSummaryMap[seg].items[item] = (segSummaryMap[seg].items[item] || 0) + q;
          segSummaryMap[seg].totalRev += r;
        }
      });

      var activeSegments = Object.keys(segSummaryMap).sort(function(a, b) {
        return segSummaryMap[b].totalRev - segSummaryMap[a].totalRev;
      });

      if (activeSegments.length === 0) {
        html += `<div style="padding:15px; border:1px solid #cbd5e1; border-radius:8px; text-align:center; color:#64748b;">${i18n.sec6Empty}</div>`;
      } else {
        html += `
          <div style="border:1px solid #cbd5e1; border-radius:8px; overflow:hidden;">
          <table style="width:100%; border-collapse:collapse; font-size:0.85rem; font-family:'Segoe UI', Arial, sans-serif;">
            <thead>
              <tr style="background-color:#f8fafc; color:#475569;">
                <th style="padding:10px; text-align:center; border-bottom:2px solid #0f766e; border-right:1px solid #cbd5e1; width:130px;">${i18n.sec6Col1}</th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #0f766e; border-right:1px solid #cbd5e1; width:110px;">${i18n.sec6Col2}</th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #0f766e; border-right:1px solid #cbd5e1; width:100px;">${i18n.sec6Col3}</th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #0f766e; border-right:1px solid #cbd5e1;">${i18n.sec6Col4}</th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #0f766e; width:120px;">${i18n.sec6Col5}</th>
              </tr>
            </thead>
            <tbody>
        `;

        activeSegments.forEach(function(segName) {
          var segObj = segSummaryMap[segName];
          var uniqueClientsCount = Object.keys(segObj.clientsMap).length;
          var totalOrdersCount = Object.keys(segObj.ordersMap).length;

          var itemListStr = Object.keys(segObj.items).map(function(pName) {
            var qty = segObj.items[pName];
            return '<div style="margin-bottom:3px;"><strong style="color:#0f766e;">' + qty + 'x</strong> ' + pName + '</div>';
          }).join("");

          html += `
            <tr style="border-top:1px solid #e2e8f0;">
              <td style="padding:9px 10px; text-align:center; border-right:1px solid #cbd5e1;"><span class="badge badge-warning">${segName}</span></td>
              <td style="padding:9px 10px; text-align:center; border-right:1px solid #cbd5e1; font-weight:bold; color:#0f766e;">${uniqueClientsCount} ${isAr ? 'عميل' : 'Clients'}</td>
              <td style="padding:9px 10px; text-align:center; border-right:1px solid #cbd5e1; font-weight:bold; color:#1e293b;">${totalOrdersCount} ${isAr ? 'طلب' : 'Orders'}</td>
              <td style="padding:9px 10px; text-align:center; border-right:1px solid #cbd5e1; color:#334155; font-size:0.8rem;">${itemListStr}</td>
              <td style="padding:9px 10px; text-align:center; font-weight:bold; color:#15803d;">${formatKM(segObj.totalRev, revUnit)}</td>
            </tr>`;
        });

        html += `</tbody></table></div>`;
      }

      html += `</div>`;
      sectionCounter++;
    }

    // --- SECTION 7: PRODUCT CATEGORY BREAKDOWN ---
    else if (secId === 7) {
      html += `
        <div style="background:#ffffff; border:1px solid #cbd5e1; border-${isAr ? 'right' : 'left'}:6px solid #0284c7; border-radius:12px; padding:20px; margin-bottom:30px; box-shadow:0 4px 6px rgba(0,0,0,0.03);">
          <div style="background-color:#0284c7; color:#ffffff; padding:10px 16px; border-radius:8px; font-weight:bold; font-size:1.05rem; margin-bottom:15px;">
            ${sectionCounter}. ${i18n.sec7Title}
          </div>
      `;

      var catMap = {};
      var grandTotalCat = { currV: 0, prevV: 0, currQ: 0, prevQ: 0 };

      Object.keys(rcData).forEach(function(dStr) {
        var dMs = new Date(dStr).getTime();
        var isCurr = (dMs >= currStartMs && dMs <= currEndMs);
        var isPrev = (dMs >= baseStartMs && dMs <= baseEndMs);

        if (isCurr || isPrev) {
          var segs = rcData[dStr];
          Object.keys(segs).forEach(function(s) {
            if (!isMatchSeg(s)) return;
            Object.keys(segs[s]).forEach(function(m) {
              Object.keys(segs[s][m]).forEach(function(b) {
                Object.keys(segs[s][m][b]).forEach(function(rawCat) {
                  var catName = categoryDict[rawCat] || rawCat;
                  if (!catMap[catName]) catMap[catName] = { currV: 0, prevV: 0, currQ: 0, prevQ: 0 };

                  Object.keys(segs[s][m][b][rawCat]).forEach(function(t) {
                    Object.keys(segs[s][m][b][rawCat][t]).forEach(function(p) {
                      var item = segs[s][m][b][rawCat][t][p];
                      var q = item.q || 0;
                      var v = item.v || 0;

                      if (isCurr) {
                        catMap[catName].currV += v;
                        catMap[catName].currQ += q;
                        grandTotalCat.currV += v;
                        grandTotalCat.currQ += q;
                      }
                      if (isPrev) {
                        catMap[catName].prevV += v;
                        catMap[catName].prevQ += q;
                        grandTotalCat.prevV += v;
                        grandTotalCat.prevQ += q;
                      }
                    });
                  });
                });
              });
            });
          });
        }
      });

      var catKeys = Object.keys(catMap).sort(function(a, b) {
        return (catMap[b].currV || 0) - (catMap[a].currV || 0);
      });

      var renderCatCompositeCell = function(prevV, currV, prevQ, currQ, colType, isCurrentCol) {
        var diffV = currV - prevV;
        var diffQ = currQ - prevQ;

        var pctV = prevV > 0 ? ((diffV / prevV) * 100) : (currV > 0 ? 100 : 0);
        var pctQ = prevQ > 0 ? ((diffQ / prevQ) * 100) : (currQ > 0 ? 100 : 0);

        var colorV = diffV >= 0 ? "#15803d" : "#b91c1c";
        var colorQ = diffQ >= 0 ? "#15803d" : "#b91c1c";
        var bgQ = diffQ >= 0 ? "#dcfce7" : "#fee2e2";

        var signV = diffV >= 0 ? "+" : "-";
        var signQ = diffQ >= 0 ? "+" : "-";

        var revHtml = "";
        var volHtml = "";

        if (colType === "PRIOR") {
          revHtml = `<strong style="color:#475569;">${formatKM(prevV, revUnit)}</strong>`;
          volHtml = `<span style="display:inline-block; font-size:0.72rem; background:#e2e8f0; color:#475569; padding:2px 6px; border-radius:10px; margin-${isAr?'right':'left'}:5px; font-weight:600; white-space:nowrap;">${formatKM(prevQ, qtyUnit)}</span>`;

        } else if (colType === "CURRENT") {
          revHtml = `<strong style="color:#0284c7; font-size:0.92rem;">${formatKM(currV, revUnit)}</strong>`;
          volHtml = `<span style="display:inline-block; font-size:0.72rem; background:#e0f2fe; color:#0369a1; padding:2px 6px; border-radius:10px; margin-${isAr?'right':'left'}:5px; font-weight:bold; white-space:nowrap;">${formatKM(currQ, qtyUnit)}</span>`;

        } else if (colType === "VARIANCE") {
          revHtml = `<strong style="color:${colorV};">${signV}${formatKM(Math.abs(diffV), revUnit)}</strong>`;
          volHtml = `<span style="display:inline-block; font-size:0.72rem; background:${bgQ}; color:${colorQ}; padding:2px 6px; border-radius:10px; margin-${isAr?'right':'left'}:5px; font-weight:bold; white-space:nowrap;">${signQ}${formatKM(Math.abs(diffQ), qtyUnit)}</span>`;

        } else if (colType === "GROWTH") {
          revHtml = `<strong style="color:${colorV};">${signV}${Math.abs(pctV).toFixed(1)}%</strong>`;
          volHtml = `<span style="display:inline-block; font-size:0.72rem; background:${bgQ}; color:${colorQ}; padding:2px 6px; border-radius:10px; margin-${isAr?'right':'left'}:5px; font-weight:bold; white-space:nowrap;">${signQ}${Math.abs(pctQ).toFixed(1)}%</span>`;
        }

        var cellBg = isCurrentCol ? "background-color:#f0f9ff;" : "";
        return `<td style="padding:9px 10px; text-align:center; border-right:1px solid #cbd5e1; ${cellBg}">${revHtml}${volHtml}</td>`;
      };

      html += `
        <div style="margin-top:10px; margin-bottom:10px; border:1px solid #cbd5e1; border-radius:8px; overflow:hidden;">
          <table style="width:100%; border-collapse:collapse; font-size:0.85rem; font-family:'Segoe UI', Arial, sans-serif;">
            <thead>
              <tr style="background-color:#f8fafc; color:#475569;">
                <th style="padding:10px; text-align:center; border-bottom:2px solid #0284c7; border-right:1px solid #cbd5e1; min-width:160px;">${i18n.sec7Col1}</th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #0284c7; border-right:1px solid #cbd5e1; color:#64748b; background-color:#f1f5f9;">${i18n.prior} (${basePeriodLabel})</th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #0284c7; border-right:1px solid #cbd5e1; color:#ffffff; background-color:#0284c7; font-size:0.9rem;">${i18n.current} (${currPeriodLabel})</th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #0284c7; border-right:1px solid #cbd5e1; background-color:#f1f5f9;">${i18n.variance}</th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #0284c7; background-color:#f1f5f9;">${i18n.growth}</th>
              </tr>
            </thead>
            <tbody>
      `;

      if (catKeys.length === 0) {
        html += `<tr><td colspan="5" style="text-align:center; color:#64748b; padding:15px;">${i18n.sec7Empty}</td></tr>`;
      } else {
        var renderCatRow = function(name, prevV, currV, prevQ, currQ, isTotal) {
          var bgRow = isTotal ? "background-color:#e2e8f0; font-weight:bold; border-top:2px solid #0284c7;" : "border-top:1px solid #e2e8f0;";
          var labelText = isTotal 
            ? `<strong style="color:#0284c7; font-size:0.95rem;">${i18n.grandTotal}</strong>` 
            : `<strong style="color:#1e293b;">${name}</strong>`;

          return `
            <tr style="${bgRow}">
              <td style="padding:9px 10px; text-align:center; border-right:1px solid #cbd5e1;">${labelText}</td>
              ${renderCatCompositeCell(prevV, currV, prevQ, currQ, "PRIOR", false)}
              ${renderCatCompositeCell(prevV, currV, prevQ, currQ, "CURRENT", true)}
              ${renderCatCompositeCell(prevV, currV, prevQ, currQ, "VARIANCE", false)}
              ${renderCatCompositeCell(prevV, currV, prevQ, currQ, "GROWTH", false)}
            </tr>`;
        };

        html += renderCatRow(i18n.grandTotal, grandTotalCat.prevV, grandTotalCat.currV, grandTotalCat.prevQ, grandTotalCat.currQ, true);

        catKeys.forEach(function(cName) {
          var cData = catMap[cName];
          html += renderCatRow(cName, cData.prevV, cData.currV, cData.prevQ, cData.currQ, false);
        });
      }

      html += `</tbody></table></div></div>`;
      sectionCounter++;
    }

    // --- SECTION 8: MACRO PROGRESS & YTD TRACKING (SCOPE-AWARE DYNAMIC SWITCH: MTD vs QTD) ---
    else if (secId === 8) {
      // 1. Determine Scope Grain (Quarterly switches to QTD; Daily/Weekly/Monthly uses MTD)
      var isQuarterScope = (tpl.timeScope === 'Quarterly');
      var qIdx = Math.floor(currMonth / 3); // 0 (Q1), 1 (Q2), 2 (Q3), 3 (Q4)
      var qStartMonth = qIdx * 3;
      var qEndMonth = qStartMonth + 2;

      var monthNamesEn = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
      var monthNamesAr = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"];
      var activeMonthName = isAr ? monthNamesAr[currMonth] : monthNamesEn[currMonth];

      // Format Date Boundaries
      var interCurrStartDate, interCurrEndDate, interBaseStartDate, interBaseEndDate;
      if (isQuarterScope) {
        interCurrStartDate = new Date(currYear, qStartMonth, 1);
        interCurrEndDate = (tpl.scheduleDay === 'Q_END') 
          ? new Date(currYear, qEndMonth + 1, 0, 23, 59, 59) 
          : new Date(currYear, currMonth, currDate, 23, 59, 59);

        interBaseStartDate = new Date(baseYear, qStartMonth, 1);
        interBaseEndDate = (tpl.scheduleDay === 'Q_END') 
          ? new Date(baseYear, qEndMonth + 1, 0, 23, 59, 59) 
          : new Date(baseYear, currMonth, currDate, 23, 59, 59);
      } else {
        interCurrStartDate = new Date(currYear, currMonth, 1);
        interCurrEndDate = new Date(currYear, currMonth, currDate, 23, 59, 59);

        interBaseStartDate = new Date(baseYear, currMonth, 1);
        interBaseEndDate = new Date(baseYear, currMonth, currDate, 23, 59, 59);
      }

      var interCurrDateStr = fmtDate(interCurrStartDate) + " - " + fmtDate(interCurrEndDate);
      var interBaseDateStr = fmtDate(interBaseStartDate) + " - " + fmtDate(interBaseEndDate);

      var interCurrStartMs = interCurrStartDate.getTime();
      var interCurrEndMs = interCurrEndDate.getTime();
      var interBaseStartMs = interBaseStartDate.getTime();
      var interBaseEndMs = interBaseEndDate.getTime();

      // Top Anchors & Cumulative Accumulators
      var baseFullYear = { v: 0, q: 0 };
      var baseFullQtr = { v: 0, q: 0 };
      var baseFullMonth = { v: 0, q: 0 };
      var baseInter = { v: 0, q: 0 };
      var currInter = { v: 0, q: 0 };
      var baseYTD = { v: 0, q: 0 };
      var currYTD = { v: 0, q: 0 };
      var interSegMap = {};

      Object.keys(rcData).forEach(function(dStr) {
        var d = new Date(dStr);
        var y = d.getFullYear();
        var m = d.getMonth();
        var day = d.getDate();
        var dMs = d.getTime();

        var isBaseFullY = (y === baseYear);
        var isBaseFullQ = (y === baseYear && Math.floor(m / 3) === qIdx);
        var isBaseFullM = (y === baseYear && m === currMonth);

        var isBaseInter = (dMs >= interBaseStartMs && dMs <= interBaseEndMs);
        var isCurrInter = (dMs >= interCurrStartMs && dMs <= interCurrEndMs);

        var isBaseY = (y === baseYear && (m < currMonth || (m === currMonth && day <= currDate)));
        var isCurrY = (y === currYear && (m < currMonth || (m === currMonth && day <= currDate)));

        if (isBaseFullY || isCurrY || isBaseInter || isCurrInter) {
          var segs = rcData[dStr];
          Object.keys(segs).forEach(function(s) {
            if (!isMatchSeg(s)) return;
            Object.keys(segs[s]).forEach(function(mc) {
              Object.keys(segs[s][mc]).forEach(function(b) {
                Object.keys(segs[s][mc][b]).forEach(function(c) {
                  Object.keys(segs[s][mc][b][c]).forEach(function(t) {
                    Object.keys(segs[s][mc][b][c][t]).forEach(function(p) {
                      var item = segs[s][mc][b][c][t][p];
                      var qVal = item.q || 0;
                      var vVal = item.v || 0;

                      if (isBaseFullY) { baseFullYear.v += vVal; baseFullYear.q += qVal; }
                      if (isBaseFullQ) { baseFullQtr.v += vVal; baseFullQtr.q += qVal; }
                      if (isBaseFullM) { baseFullMonth.v += vVal; baseFullMonth.q += qVal; }

                      if (isBaseInter) {
                        baseInter.v += vVal; baseInter.q += qVal;
                        if (!interSegMap[s]) interSegMap[s] = { currV: 0, prevV: 0, currQ: 0, prevQ: 0 };
                        interSegMap[s].prevV += vVal;
                        interSegMap[s].prevQ += qVal;
                      }

                      if (isCurrInter) {
                        currInter.v += vVal; currInter.q += qVal;
                        if (!interSegMap[s]) interSegMap[s] = { currV: 0, prevV: 0, currQ: 0, prevQ: 0 };
                        interSegMap[s].currV += vVal;
                        interSegMap[s].currQ += qVal;
                      }

                      if (isBaseY) { baseYTD.v += vVal; baseYTD.q += qVal; }
                      if (isCurrY) { currYTD.v += vVal; currYTD.q += qVal; }
                    });
                  });
                });
              });
            });
          });
        }
      });

      // 2. Micro-Progress Bar Generator Helper (Rock-solid HTML Table based)
      var renderProgressBarLine = function(label, val, baseFullVal, unit, color, isBoldCurrent) {
        var pct = baseFullVal > 0 ? (val / baseFullVal) * 100 : 0;
        var barWidth = Math.min(Math.max(pct, 0), 100);
        var labelWeight = isBoldCurrent ? "800" : "600";
        var labelColor = isBoldCurrent ? "#003366" : "#475569";
        var barHeight = isBoldCurrent ? "8px" : "6px";

        return `
          <table style="width:100%; border-collapse:collapse; margin-bottom:5px;">
            <tr>
              <td style="font-size:0.7rem; font-weight:${labelWeight}; color:${labelColor}; padding:0; text-align:${align};">
                ${label}: <strong style="color:${isBoldCurrent ? '#003366' : '#334155'}; font-size:${isBoldCurrent ? '0.78rem' : '0.7rem'};">${formatKM(val, unit)}</strong>
              </td>
              <td style="font-size:0.7rem; font-weight:800; color:${color}; padding:0; text-align:${isAr ? 'left' : 'right'}; width:55px;">
                ${pct.toFixed(1)}%
              </td>
            </tr>
            <tr>
              <td colspan="2" style="padding-top:2px; padding-bottom:3px;">
                <div style="background-color:#e2e8f0; border-radius:4px; height:${barHeight}; width:100%; overflow:hidden;">
                  <div style="background-color:${color}; width:${barWidth}%; height:${barHeight}; border-radius:4px;"></div>
                </div>
              </td>
            </tr>
          </table>
        `;
      };

      // 3. Composite Metric Cell Renderer (Prominent Current vs Baseline)
      var renderMacroCell = function(currV, currQ, isProminentCurrent) {
        var bgStyle = isProminentCurrent ? "background-color:#eff6ff; border:1px solid #bfdbfe;" : "";
        var revColor = isProminentCurrent ? "#003366" : "#475569";
        var revSize = isProminentCurrent ? "1.05rem" : "0.92rem";
        var qtyBg = isProminentCurrent ? "#003366" : "#cbd5e1";
        var qtyColor = isProminentCurrent ? "#ffffff" : "#0f172a";

        return `
          <td style="padding:10px; text-align:center; border-right:1px solid #cbd5e1; ${bgStyle}">
            <strong style="color:${revColor}; font-size:${revSize}; font-weight:800;">${formatKM(currV, revUnit)}</strong>
            <span style="display:inline-block; font-size:0.72rem; background:${qtyBg}; color:${qtyColor}; padding:2px 7px; border-radius:10px; margin-${isAr?'right':'left'}:6px; font-weight:bold; white-space:nowrap;">
              ${formatKM(currQ, qtyUnit)}
            </span>
          </td>
        `;
      };

      // 4. Growth Cell Renderer (Rev % dominant + Qty % chip with independent colors)
      var renderMacroGrowthCell = function(diffV, pctV, diffQ, pctQ) {
        var colorV = diffV >= 0 ? "#15803d" : "#b91c1c";
        var colorQ = diffQ >= 0 ? "#15803d" : "#b91c1c";
        var bgQ = diffQ >= 0 ? "#dcfce7" : "#fee2e2";
        var signV = diffV >= 0 ? "+" : "-";
        var signQ = diffQ >= 0 ? "+" : "-";

        return `
          <td style="padding:10px; text-align:center;">
            <strong style="color:${colorV}; font-size:0.95rem; font-weight:800;">${signV}${Math.abs(pctV).toFixed(1)}%</strong>
            <span style="display:inline-block; font-size:0.72rem; background:${bgQ}; color:${colorQ}; padding:2px 6px; border-radius:10px; margin-${isAr?'right':'left'}:5px; font-weight:bold; white-space:nowrap;">
              ${signQ}${Math.abs(pctQ).toFixed(1)}%
            </span>
          </td>
        `;
      };

      // Dynamic Section Titles & Labels
      var sec8Title = isAr ? "متابعة الأداء الكلي والتراكمي (MTD / QTD / YTD)" : "Macro Progress & YTD Tracking";
      var fullBaseYearLabel = isAr ? `الهدف: سنة ${baseYear} كاملة` : `Target: Full Base Year (${baseYear})`;
      var fullBaseQtrLabel = isAr ? `الهدف: الربع الحالي كامل لسنة الأساس (Q${qIdx+1} - ${baseYear})` : `Target: Full Base Q${qIdx+1} (${baseYear})`;
      var fullBaseMonthLabel = isAr ? `الهدف: شهر ${activeMonthName} كامل لسنة الأساس (${baseYear})` : `Target: Full Base Month (${activeMonthName} ${baseYear})`;

      // Dynamic Intermediate Routing Variables
      var targetInter = isQuarterScope ? baseFullQtr : baseFullMonth;
      var targetInterLabel = isQuarterScope 
        ? (isAr ? `الهدف: ربع كامل (Q${qIdx+1} - ${baseYear})` : `Target: Full Quarter (Q${qIdx+1} - ${baseYear})`)
        : (isAr ? "الهدف: شهر كامل" : "Target: Full Month");

      var interCode = isQuarterScope ? "QTD" : "MTD";
      var interScopeTitle = isQuarterScope
        ? (isAr ? `متابعة الربع الحالي حتي تاريخه (QTD: ${interCurrDateStr})` : `CURRENT QUARTER MACRO-TRACKING (QTD: ${interCurrDateStr})`)
        : (isAr ? `متابعة الشهر الحالي حتي تاريخه (MTD: ${interCurrDateStr})` : `CURRENT MONTH MACRO-TRACKING (MTD: ${interCurrDateStr})`);

      var interCurrSalesColHeader = isQuarterScope
        ? (isAr ? `مبيعات الربع حتي تاريخه (${currYear})` : `Current QTD Sales (${currYear})`)
        : (isAr ? `مبيعات الشهر حتي تاريخه (${currYear})` : `Current MTD Sales (${currYear})`);

      var interBaseSalesColHeader = isQuarterScope
        ? (isAr ? `مبيعات نفس الربع سنة الأساس (${baseYear})` : `Base QTD Sales (${baseYear})`)
        : (isAr ? `مبيعات نفس الشهر سنة الأساس (${baseYear})` : `Base MTD Sales (${baseYear})`);

      var interCurrPacingLineLabel = isQuarterScope
        ? (isAr ? `الربع حتي تاريخه ${currYear}` : `Quarter to-date ${currYear}`)
        : (isAr ? `الشهر حتي تاريخه ${currYear}` : `Month to-date ${currYear}`);

      var interBasePacingLineLabel = isQuarterScope
        ? (isAr ? `الربع حتي تاريخه ${baseYear}` : `Quarter to-date ${baseYear}`)
        : (isAr ? `الشهر حتي تاريخه ${baseYear}` : `Month to-date ${baseYear}`);

      html += `
        <div style="background:#ffffff; border:1px solid #cbd5e1; border-${isAr ? 'right' : 'left'}:6px solid #475569; border-radius:12px; padding:20px; margin-bottom:30px; box-shadow:0 4px 6px rgba(0,0,0,0.03);">
          <div style="background-color:#334155; color:#ffffff; padding:10px 16px; border-radius:8px; font-weight:bold; font-size:1.05rem; margin-bottom:15px;">
            ${sectionCounter}. ${sec8Title}
          </div>

          <!-- TOP ANCHORED BASELINE TARGETS (3 HORIZONTAL CARDS) -->
          <div style="background-color:#f8fafc; border:1px solid #cbd5e1; border-radius:10px; padding:12px; margin-bottom:20px;">
            <div style="font-weight:bold; color:#003366; font-size:0.75rem; text-align:center; margin-bottom:10px; text-transform:uppercase; letter-spacing:0.5px;">
              <i class="fas fa-bullseye me-1"></i>${isAr ? `الهدف: مؤشرات سنة الأساس المرجعية (${baseYear})` : `Anchored Baseline Targets (${baseYear})`}
            </div>
            <table style="width:100%; border-collapse:separate; border-spacing:8px 0;">
              <tr>
                <!-- CARD 1: FULL BASE YEAR -->
                <td style="width:33.33%; background:#ffffff; border:1px solid #cbd5e1; border-radius:8px; padding:10px; text-align:center; vertical-align:top;">
                  <div style="font-size:0.65rem; font-weight:800; color:#475569; text-transform:uppercase;">${fullBaseYearLabel}</div>
                  <div style="font-size:1rem; font-weight:800; color:#003366; margin-top:3px;">${formatKM(baseFullYear.v, revUnit)}</div>
                  <div style="margin-top:4px;">
                    <span style="font-size:0.72rem; background:#f1f5f9; color:#475569; padding:2px 8px; border-radius:10px; font-weight:bold;">${formatKM(baseFullYear.q, qtyUnit)}</span>
                  </div>
                </td>

                <!-- CARD 2: FULL BASE QUARTER -->
                <td style="width:33.33%; background:#ffffff; border:1px solid #cbd5e1; border-radius:8px; padding:10px; text-align:center; vertical-align:top;">
                  <div style="font-size:0.65rem; font-weight:800; color:#475569; text-transform:uppercase;">${fullBaseQtrLabel}</div>
                  <div style="font-size:1rem; font-weight:800; color:#003366; margin-top:3px;">${formatKM(baseFullQtr.v, revUnit)}</div>
                  <div style="margin-top:4px;">
                    <span style="font-size:0.72rem; background:#f1f5f9; color:#475569; padding:2px 8px; border-radius:10px; font-weight:bold;">${formatKM(baseFullQtr.q, qtyUnit)}</span>
                  </div>
                </td>

                <!-- CARD 3: FULL BASE MONTH -->
                <td style="width:33.33%; background:#ffffff; border:1px solid #cbd5e1; border-radius:8px; padding:10px; text-align:center; vertical-align:top;">
                  <div style="font-size:0.65rem; font-weight:800; color:#475569; text-transform:uppercase;">${fullBaseMonthLabel}</div>
                  <div style="font-size:1rem; font-weight:800; color:#003366; margin-top:3px;">${formatKM(baseFullMonth.v, revUnit)}</div>
                  <div style="margin-top:4px;">
                    <span style="font-size:0.72rem; background:#f1f5f9; color:#475569; padding:2px 8px; border-radius:10px; font-weight:bold;">${formatKM(baseFullMonth.q, qtyUnit)}</span>
                  </div>
                </td>
              </tr>
            </table>
          </div>
      `;

      // --- TABLE 1: DYNAMIC INTERMEDIATE TRACKING (MTD or QTD) ---
      var diffInterV = currInter.v - baseInter.v;
      var pctInterV = baseInter.v > 0 ? ((diffInterV / baseInter.v) * 100) : (currInter.v > 0 ? 100 : 0);
      var diffInterQ = currInter.q - baseInter.q;
      var pctInterQ = baseInter.q > 0 ? ((diffInterQ / baseInter.q) * 100) : (currInter.q > 0 ? 100 : 0);

      var currInterColorQ = currInter.q >= baseInter.q ? "#15803d" : "#b91c1c";
      var currInterColorV = currInter.v >= baseInter.v ? "#15803d" : "#b91c1c";

      html += `
        <div style="margin-top:10px; margin-bottom:25px; border:1px solid #cbd5e1; border-radius:8px; overflow:hidden;">
          <div style="background-color:#0f766e; color:#ffffff; padding:8px 14px; font-weight:bold; font-size:0.85rem; text-align:center;">
            ${interScopeTitle}
          </div>
          <table style="width:100%; border-collapse:collapse; font-size:0.85rem; font-family:'Segoe UI', Arial, sans-serif;">
            <thead>
              <tr style="background-color:#f8fafc; color:#475569;">
                <th style="padding:10px; text-align:center; border-bottom:2px solid #0f766e; border-right:1px solid #cbd5e1; width:22%; background-color:#e0f2fe; color:#0369a1; font-weight:800;">
                  ${interCurrSalesColHeader}
                </th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #0f766e; border-right:1px solid #cbd5e1; width:22%;">
                  ${interBaseSalesColHeader}
                </th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #0f766e; border-right:1px solid #cbd5e1; width:23%;">
                  ${isAr ? `متابعة الكميات (مقارنة بالهدف)` : `Volume Pacing vs Target`}
                </th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #0f766e; border-right:1px solid #cbd5e1; width:23%;">
                  ${isAr ? `متابعة الايرادات (مقارنة بالهدف)` : `Revenue Pacing vs Target`}
                </th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #0f766e; width:10%;">
                  ${isAr ? `النمو حتي تاريخه (${interCode})` : `Growth % (${interCode})`}
                </th>
              </tr>
            </thead>
            <tbody>
              <tr style="background-color:#ffffff;">
                ${renderMacroCell(currInter.v, currInter.q, true)}
                ${renderMacroCell(baseInter.v, baseInter.q, false)}
                <td style="padding:8px 10px; text-align:center; border-right:1px solid #cbd5e1;">
                  ${renderProgressBarLine(targetInterLabel, targetInter.q, targetInter.q, qtyUnit, "#64748b", false)}
                  ${renderProgressBarLine(interBasePacingLineLabel, baseInter.q, targetInter.q, qtyUnit, "#0284c7", false)}
                  ${renderProgressBarLine(interCurrPacingLineLabel, currInter.q, targetInter.q, qtyUnit, currInterColorQ, true)}
                </td>
                <td style="padding:8px 10px; text-align:center; border-right:1px solid #cbd5e1;">
                  ${renderProgressBarLine(targetInterLabel, targetInter.v, targetInter.v, revUnit, "#64748b", false)}
                  ${renderProgressBarLine(interBasePacingLineLabel, baseInter.v, targetInter.v, revUnit, "#0284c7", false)}
                  ${renderProgressBarLine(interCurrPacingLineLabel, currInter.v, targetInter.v, revUnit, currInterColorV, true)}
                </td>
                ${renderMacroGrowthCell(diffInterV, pctInterV, diffInterQ, pctInterQ)}
              </tr>
            </tbody>
          </table>
        </div>
      `;

      // --- TABLE 2: DYNAMIC SEGMENT PERFORMANCE BREAKDOWN (MTD or QTD) ---
      var interSegKeys = Object.keys(interSegMap).sort(function(a, b) {
        return (interSegMap[b].currV || 0) - (interSegMap[a].currV || 0);
      });

      var renderInterSegCell = function(prevV, currV, prevQ, currQ, colType, isCurrentCol) {
        var diffV = currV - prevV;
        var diffQ = currQ - prevQ;

        var pctV = prevV > 0 ? ((diffV / prevV) * 100) : (currV > 0 ? 100 : 0);
        var pctQ = prevQ > 0 ? ((diffQ / prevQ) * 100) : (currQ > 0 ? 100 : 0);

        var colorV = diffV >= 0 ? "#15803d" : "#b91c1c";
        var colorQ = diffQ >= 0 ? "#15803d" : "#b91c1c";
        var bgQ = diffQ >= 0 ? "#dcfce7" : "#fee2e2";

        var signV = diffV >= 0 ? "+" : "-";
        var signQ = diffQ >= 0 ? "+" : "-";

        var revHtml = "";
        var volHtml = "";

        if (colType === "PRIOR") {
          revHtml = `<strong style="color:#475569;">${formatKM(prevV, revUnit)}</strong>`;
          volHtml = `<span style="display:inline-block; font-size:0.72rem; background:#e2e8f0; color:#475569; padding:2px 6px; border-radius:10px; margin-${isAr?'right':'left'}:5px; font-weight:600; white-space:nowrap;">${formatKM(prevQ, qtyUnit)}</span>`;
        } else if (colType === "CURRENT") {
          revHtml = `<strong style="color:#003366; font-size:0.95rem; font-weight:800;">${formatKM(currV, revUnit)}</strong>`;
          volHtml = `<span style="display:inline-block; font-size:0.72rem; background:#003366; color:#ffffff; padding:2px 7px; border-radius:10px; margin-${isAr?'right':'left'}:5px; font-weight:bold; white-space:nowrap;">${formatKM(currQ, qtyUnit)}</span>`;
        } else if (colType === "VARIANCE") {
          revHtml = `<strong style="color:${colorV}; font-weight:800;">${signV}${formatKM(Math.abs(diffV), revUnit)}</strong>`;
          volHtml = `<span style="display:inline-block; font-size:0.72rem; background:${bgQ}; color:${colorQ}; padding:2px 6px; border-radius:10px; margin-${isAr?'right':'left'}:5px; font-weight:bold; white-space:nowrap;">${signQ}${formatKM(Math.abs(diffQ), qtyUnit)}</span>`;
        } else if (colType === "GROWTH") {
          revHtml = `<strong style="color:${colorV}; font-weight:800;">${signV}${Math.abs(pctV).toFixed(1)}%</strong>`;
          volHtml = `<span style="display:inline-block; font-size:0.72rem; background:${bgQ}; color:${colorQ}; padding:2px 6px; border-radius:10px; margin-${isAr?'right':'left'}:5px; font-weight:bold; white-space:nowrap;">${signQ}${Math.abs(pctQ).toFixed(1)}%</span>`;
        }

        var cellBg = isCurrentCol ? "background-color:#eff6ff;" : "";
        return `<td style="padding:9px 10px; text-align:center; border-right:1px solid #cbd5e1; ${cellBg}">${revHtml}${volHtml}</td>`;
      };

      var segTableHeading = isQuarterScope
        ? (isAr ? `تحليل أداء القطاعات خلال الربع حتي تاريخه (QTD)` : `QTD SEGMENT PERFORMANCE BREAKDOWN`)
        : (isAr ? `تحليل أداء القطاعات خلال الشهر حتي تاريخه (MTD)` : `MTD SEGMENT PERFORMANCE BREAKDOWN`);

      html += `
        <div style="margin-top:10px; margin-bottom:25px; border:1px solid #cbd5e1; border-radius:8px; overflow:hidden;">
          <div style="background-color:#475569; color:#ffffff; padding:8px 14px; font-weight:bold; font-size:0.85rem; text-align:center;">
            ${segTableHeading}
          </div>
          <table style="width:100%; border-collapse:collapse; font-size:0.85rem; font-family:'Segoe UI', Arial, sans-serif;">
            <thead>
              <tr style="background-color:#f8fafc; color:#475569;">
                <th style="padding:10px; text-align:center; border-bottom:2px solid #475569; border-right:1px solid #cbd5e1; min-width:140px;">${i18n.sec1Col1}</th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #475569; border-right:1px solid #cbd5e1; color:#64748b; background-color:#f1f5f9;">${i18n.prior} (${interBaseDateStr})</th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #475569; border-right:1px solid #cbd5e1; color:#ffffff; background-color:#003366; font-size:0.92rem; font-weight:800;">${i18n.current} (${interCurrDateStr})</th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #475569; border-right:1px solid #cbd5e1; background-color:#f1f5f9;">${i18n.variance}</th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #475569; background-color:#f1f5f9;">${i18n.growth}</th>
              </tr>
            </thead>
            <tbody>
      `;

      if (interSegKeys.length === 0) {
        html += `<tr><td colspan="5" style="text-align:center; color:#64748b; padding:15px;">${isAr ? "لا توجد بيانات مسجلة للقطاعات خلال هذا النطاق." : "No segment data recorded for this scope."}</td></tr>`;
      } else {
        var renderInterSegRow = function(name, prevV, currV, prevQ, currQ, isTotal) {
          var bgRow = isTotal ? "background-color:#e2e8f0; font-weight:bold; border-top:2px solid #475569;" : "border-top:1px solid #e2e8f0;";
          var labelText = isTotal 
            ? `<strong style="color:#003366; font-size:0.95rem;">${i18n.grandTotal}</strong>` 
            : `<strong style="color:#1e293b;">${name}</strong>`;

          return `
            <tr style="${bgRow}">
              <td style="padding:9px 10px; text-align:center; border-right:1px solid #cbd5e1;">${labelText}</td>
              ${renderInterSegCell(prevV, currV, prevQ, currQ, "PRIOR", false)}
              ${renderInterSegCell(prevV, currV, prevQ, currQ, "CURRENT", true)}
              ${renderInterSegCell(prevV, currV, prevQ, currQ, "VARIANCE", false)}
              ${renderInterSegCell(prevV, currV, prevQ, currQ, "GROWTH", false)}
            </tr>`;
        };

        html += renderInterSegRow(i18n.grandTotal, baseInter.v, currInter.v, baseInter.q, currInter.q, true);

        interSegKeys.forEach(function(sName) {
          var sData = interSegMap[sName];
          html += renderInterSegRow(sName, sData.prevV, sData.currV, sData.prevQ, sData.currQ, false);
        });
      }

      html += `</tbody></table></div>`;

      // --- TABLE 3: YEAR-TO-DATE MACRO-TRACKING (YTD - AT THE BOTTOM) ---
      var diffYtdV = currYTD.v - baseYTD.v;
      var pctYtdV = baseYTD.v > 0 ? ((diffYtdV / baseYTD.v) * 100) : (currYTD.v > 0 ? 100 : 0);
      var diffYtdQ = currYTD.q - baseYTD.q;
      var pctYtdQ = baseYTD.q > 0 ? ((diffYtdQ / baseYTD.q) * 100) : (currYTD.q > 0 ? 100 : 0);

      var currYtdColorQ = currYTD.q >= baseYTD.q ? "#15803d" : "#b91c1c";
      var currYtdColorV = currYTD.v >= baseYTD.v ? "#15803d" : "#b91c1c";

      var ytdDateStr = "01/01/" + currYear + " - " + fmtDate(new Date(currYear, currMonth, currDate));

      html += `
        <div style="margin-top:10px; margin-bottom:10px; border:1px solid #cbd5e1; border-radius:8px; overflow:hidden;">
          <div style="background-color:#003366; color:#ffffff; padding:8px 14px; font-weight:bold; font-size:0.85rem; text-align:center;">
            ${isAr ? `متابعة الأداء التراكمي للسنة حتي تاريخه (YTD: ${ytdDateStr})` : `YEAR-TO-DATE MACRO-TRACKING (YTD: ${ytdDateStr})`}
          </div>
          <table style="width:100%; border-collapse:collapse; font-size:0.85rem; font-family:'Segoe UI', Arial, sans-serif;">
            <thead>
              <tr style="background-color:#f8fafc; color:#475569;">
                <th style="padding:10px; text-align:center; border-bottom:2px solid #003366; border-right:1px solid #cbd5e1; width:22%; background-color:#e0f2fe; color:#0369a1; font-weight:800;">
                  ${isAr ? `المبيعات التراكمية للسنة حتي تاريخه (${currYear})` : `Current YTD Sales (${currYear})`}
                </th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #003366; border-right:1px solid #cbd5e1; width:22%;">
                  ${isAr ? `مبيعات نفس الفترة للسنة حتي تاريخه (${baseYear})` : `Base YTD Sales (${baseYear})`}
                </th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #003366; border-right:1px solid #cbd5e1; width:23%;">
                  ${isAr ? `متابعة الكميات (مقارنة بالهدف)` : `Volume Pacing vs Target`}
                </th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #003366; border-right:1px solid #cbd5e1; width:23%;">
                  ${isAr ? `متابعة الايرادات (مقارنة بالهدف)` : `Revenue Pacing vs Target`}
                </th>
                <th style="padding:10px; text-align:center; border-bottom:2px solid #003366; width:10%;">
                  ${isAr ? `النمو التراكمي حتي تاريخه (YTD)` : `Growth % (YTD)`}
                </th>
              </tr>
            </thead>
            <tbody>
              <tr style="background-color:#ffffff;">
                ${renderMacroCell(currYTD.v, currYTD.q, true)}
                ${renderMacroCell(baseYTD.v, baseYTD.q, false)}
                <td style="padding:8px 10px; text-align:center; border-right:1px solid #cbd5e1;">
                  ${renderProgressBarLine(isAr ? "الهدف: سنة كاملة" : "Target: Full Year", baseFullYear.q, baseFullYear.q, qtyUnit, "#64748b", false)}
                  ${renderProgressBarLine(isAr ? `السنة حتي تاريخه ${baseYear}` : `Year to-date ${baseYear}`, baseYTD.q, baseFullYear.q, qtyUnit, "#0284c7", false)}
                  ${renderProgressBarLine(isAr ? `السنة حتي تاريخه ${currYear}` : `Year to-date ${currYear}`, currYTD.q, baseFullYear.q, qtyUnit, currYtdColorQ, true)}
                </td>
                <td style="padding:8px 10px; text-align:center; border-right:1px solid #cbd5e1;">
                  ${renderProgressBarLine(isAr ? "الهدف: سنة كاملة" : "Target: Full Year", baseFullYear.v, baseFullYear.v, revUnit, "#64748b", false)}
                  ${renderProgressBarLine(isAr ? `السنة حتي تاريخه ${baseYear}` : `Year to-date ${baseYear}`, baseYTD.v, baseFullYear.v, revUnit, "#0284c7", false)}
                  ${renderProgressBarLine(isAr ? `السنة حتي تاريخه ${currYear}` : `Year to-date ${currYear}`, currYTD.v, baseFullYear.v, revUnit, currYtdColorV, true)}
                </td>
                ${renderMacroGrowthCell(diffYtdV, pctYtdV, diffYtdQ, pctYtdQ)}
              </tr>
            </tbody>
          </table>
        </div>
      `;

      html += `</div>`;
      sectionCounter++;
    }

  });

  var emailOptions = { 
    htmlBody: html,
    name: tpl.senderName || "Sales Performance Engine" // Dynamic Sender Display Name
  };
  
  if (!recipientOverride && tpl.ccRecipients) {
    emailOptions.cc = tpl.ccRecipients;
  }

  var subjectPrefix = recipientOverride ? "[TEST MODE] " : "";
  GmailApp.sendEmail(to, subjectPrefix + tpl.templateName + " (" + currPeriodLabel + ")", "Please view in HTML-compatible reader.", emailOptions);

  // Update Last_Sent timestamp on production (scheduled) runs
  if (!recipientOverride && tpl.templateId) {
    updateTemplateLastSent(tpl.templateId);
  }

  return "Digest email successfully sent to " + to;
}

// ==========================================
// ROBUST CLIENT-LEVEL CADENCE ENGINE (FULL HISTORY)
// ==========================================

// Helper: Safely parses dates across multiple formats (ISO, Timestamps, DD/MM/YYYY)
function parseDateMs(dateVal) {
  if (!dateVal) return null;
  if (typeof dateVal === 'number') return dateVal;
  if (dateVal instanceof Date) return dateVal.getTime();
  
  var str = String(dateVal).trim();
  
  // Handle DD/MM/YYYY or DD-MM-YYYY
  if (str.indexOf('/') !== -1 || str.indexOf('-') !== -1) {
    var parts = str.split(/[\/\-]/);
    if (parts.length === 3) {
      if (parseInt(parts[0]) > 12 || parts[2].length === 4) {
        var day = parseInt(parts[0], 10);
        var month = parseInt(parts[1], 10) - 1;
        var year = parseInt(parts[2], 10);
        var d = new Date(year, month, day);
        if (!isNaN(d.getTime())) return d.getTime();
      }
    }
  }
  
  var standardParsed = new Date(str).getTime();
  return isNaN(standardParsed) ? null : standardParsed;
}

function processClientCadence(data) {
  if (!data || !Array.isArray(data)) return [];

  // Normalize TODAY to midnight for clean day integer math
  var now = new Date();
  now.setHours(0, 0, 0, 0);
  var todayMs = now.getTime();
  var msPerDay = 1000 * 60 * 60 * 24;
  var fourMonthsAgoMs = todayMs - (120 * msPerDay);

  // 1. Group overall transactions by client and order date
  var grouped = {};
  data.forEach(function(row) {
    var client = row.client || row.Client;
    if (!client) return;

    var rawDateMs = parseDateMs(row.date || row.Date || row.dateMs);
    if (!rawDateMs) return;

    var d = new Date(rawDateMs);
    d.setHours(0, 0, 0, 0);
    var dateKey = d.getTime();

    var segment = String(row.segment || row.Segment || 'Unassigned').trim();

    if (!grouped[client]) grouped[client] = {};
    if (!grouped[client][dateKey]) {
      grouped[client][dateKey] = {
        date: dateKey,
        totalQty: 0,
        segment: segment,
        region: row.region || '',
        contact: row.contactInfo || ''
      };
    }

    grouped[client][dateKey].totalQty += (Number(row.qty || row.Qty || row.quantity) || 0);
  });

  var alerts = [];

  // 2. Evaluate reorder cadence for each client
  for (var client in grouped) {
    var dateKeys = Object.keys(grouped[client]);
    var orders = dateKeys.map(function(k) { return grouped[client][k]; }).sort(function(a, b) { return a.date - b.date; });
    
    if (orders.length === 0) continue;

    var lastOrder = orders[orders.length - 1];

    // RULE: ONLY the LAST order has to fall within the last 120 days
    if (lastOrder.date < fourMonthsAgoMs) continue;

    var avgCycleDays = 0;

    if (orders.length >= 2) {
      // Calculate average cycle using ALL historical orders
      var firstOrder = orders[0];
      var totalDaysElapsed = (lastOrder.date - firstOrder.date) / msPerDay;
      var intervalsCount = orders.length - 1;
      
      if (totalDaysElapsed > 0 && intervalsCount > 0) {
        avgCycleDays = totalDaysElapsed / intervalsCount;
      }
    }

    // Baseline fallback for 1-order clients
    if (avgCycleDays <= 0) {
      avgCycleDays = 30;
    }

    // Calculate Average Order Quantity across all historical orders
    var totalQtySum = 0;
    orders.forEach(function(o) { totalQtySum += o.totalQty; });
    var avgOrderQty = totalQtySum / orders.length;

    // Expected Next Order Date & Overdue Calculation
    var expectedNextDateMs = lastOrder.date + (avgCycleDays * msPerDay);
    var daysLate = (todayMs - expectedNextDateMs) / msPerDay;

    // Filter for clients overdue by 2 or more days
    if (daysLate >= 2) {
      alerts.push({
        client: client,
        segment: lastOrder.segment,
        region: lastOrder.region,
        contact: lastOrder.contact,
        lastOrderDate: lastOrder.date,
        lastOrderQty: Math.round(lastOrder.totalQty),
        avgOrderQty: Math.round(avgOrderQty),
        expectedInterval: Math.round(avgCycleDays),
        daysLate: Math.floor(daysLate)
      });
    }
  }

  // 3. Sort alert list by Average Order Quantity (Descending)
  alerts.sort(function(a, b) {
    return b.avgOrderQty - a.avgOrderQty;
  });

  return alerts;
}

// ==========================================
// EMAIL-ONLY CADENCE ENGINE (QUANTITY-WEIGHTED, LAST 4 ORDERS MAX)
// ==========================================
function processEmailClientCadence(data) {
  if (!data || !Array.isArray(data)) return [];

  var now = new Date();
  now.setHours(0, 0, 0, 0);
  var todayMs = now.getTime();
  var msPerDay = 1000 * 60 * 60 * 24;
  var fourMonthsAgoMs = todayMs - (120 * msPerDay);

  // 1. Group overall transactions by client and order date
  var grouped = {};
  data.forEach(function(row) {
    var client = row.client || row.Client;
    if (!client) return;

    var rawDateMs = parseDateMs(row.date || row.Date || row.dateMs);
    if (!rawDateMs) return;

    var d = new Date(rawDateMs);
    d.setHours(0, 0, 0, 0);
    var dateKey = d.getTime();

    var segment = String(row.segment || row.Segment || 'Unassigned').trim();

    if (!grouped[client]) grouped[client] = {};
    if (!grouped[client][dateKey]) {
      grouped[client][dateKey] = {
        date: dateKey,
        totalQty: 0,
        segment: segment,
        region: row.region || '',
        contact: row.contactInfo || ''
      };
    }

    grouped[client][dateKey].totalQty += (Number(row.qty || row.Qty || row.quantity) || 0);
  });

  var alerts = [];

  // 2. Evaluate reorder cadence using ONLY the last 4 orders max
  for (var client in grouped) {
    var dateKeys = Object.keys(grouped[client]);
    var allOrders = dateKeys.map(function(k) { return grouped[client][k]; }).sort(function(a, b) { return a.date - b.date; });

    if (allOrders.length === 0) continue;

    var lastOrder = allOrders[allOrders.length - 1];

    // RULE: Last order must fall within the last 120 days
    if (lastOrder.date < fourMonthsAgoMs) continue;

    // SLICE TO MAX LAST 4 ORDERS (.slice(-4))
    var recentOrders = allOrders.slice(-4);
    var expectedCycleDays = 30; // Fallback default

    if (recentOrders.length >= 2) {
      var firstRecent = recentOrders[0];
      var latestRecent = recentOrders[recentOrders.length - 1];
      var totalDaysElapsed = (latestRecent.date - firstRecent.date) / msPerDay;

      // Sum quantities consumed across intervals prior to the latest order
      var qConsumed = 0;
      for (var j = 0; j < recentOrders.length - 1; j++) {
        qConsumed += recentOrders[j].totalQty;
      }

      if (totalDaysElapsed > 0 && qConsumed > 0) {
        var dailyRate = qConsumed / totalDaysElapsed; // Units consumed per day
        if (dailyRate > 0) {
          // Project expected lifespan based on LATEST order quantity
          expectedCycleDays = latestRecent.totalQty / dailyRate;
        }
      }
    }

    // Safety guardrails: Keep expected cycle within reasonable limits (3 to 180 days)
    expectedCycleDays = Math.max(3, Math.min(180, expectedCycleDays));

    // Calculate Average Order Quantity across the recent orders
    var recentQtySum = 0;
    recentOrders.forEach(function(o) { recentQtySum += o.totalQty; });
    var avgOrderQty = recentQtySum / recentOrders.length;

    // Expected Next Order Date & Overdue Calculation
    var expectedNextDateMs = lastOrder.date + (expectedCycleDays * msPerDay);
    var daysLate = (todayMs - expectedNextDateMs) / msPerDay;

    // Filter for clients overdue by 2 or more days
    if (daysLate >= 2) {
      alerts.push({
        client: client,
        segment: lastOrder.segment,
        region: lastOrder.region,
        contact: lastOrder.contact,
        lastOrderDate: lastOrder.date,
        lastOrderQty: Math.round(lastOrder.totalQty),
        avgOrderQty: Math.round(avgOrderQty),
        expectedInterval: Math.round(expectedCycleDays),
        daysLate: Math.floor(daysLate)
      });
    }
  }

  // Sort alert list by Average Order Quantity (Descending)
  alerts.sort(function(a, b) { return b.avgOrderQty - a.avgOrderQty; });

  return alerts;
}
