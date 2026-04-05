/**
 * ============================================================================
 * ERP COSTING ENGINE - PHASE 1 & 2
 * Two-Tier FIFO (First-In-First-Out) with Nested Bill of Materials (BOM)
 * ============================================================================
 */

// --- PHASE 1: DATA EXTRACTION & MAPPING ---

function buildCostEngineMaps() {
  const ss = SpreadsheetApp.openById("1NTLovSrQLtFfebXrSOuWitB29VUV4ifLHmc-Rt_MxWo");
  
  return {
    products: mapProducts(ss.getSheetByName("Products")),
    recipes: mapRecipes(ss.getSheetByName("RECIPE")),
    plans: mapPlans(ss.getSheetByName("PLANS")),
    manufacturing: mapManufacturing(ss.getSheetByName("MANUFACTURING")),
    fifoQueues: buildFIFOQueues(ss.getSheetByName("Purchased items")),
    // NEW: Load the Market Prices
    marketPrices: buildMarketPrices(ss.getSheetByName("Purchased items"), ss.getSheetByName("Price Updates"))
  };
}

function mapProducts(sheet) {
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  let products = {}; 
  for (let i = 1; i < data.length; i++) {
    let concise = String(data[i][1]).trim(); 
    let full = String(data[i][5]).trim();    
    if (concise) {
      products[concise] = full;
    }
  }
  return products;
}

function mapRecipes(sheet) {
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  let recipes = {}; 
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const recipeId = String(row[0]).trim();
    if (!recipeId) continue;
    
    recipes[recipeId] = { materials: {} };
    
    for (let col = 6; col < headers.length; col++) {
      let genMatName = String(headers[col]).trim();
      let qty = Number(row[col]) || 0;
      if (genMatName && qty > 0) {
        recipes[recipeId].materials[genMatName] = qty;
      }
    }
  }
  return recipes;
}

function mapPlans(sheet) {
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  let plans = {}; 
  
  for (let i = 2; i < data.length; i++) {
    const row = data[i];
    if (String(row[0]).toUpperCase() !== "TRUE") continue; 
    
    const planId = String(row[1]).trim();
    if (!planId) continue;
    
    plans[planId] = { materials: {} };
    
    for (let col = 6; col < row.length; col += 2) {
      let genMatName = String(row[col]).trim();
      let specificId = String(row[col + 1]).trim();
      if (genMatName && specificId) {
        plans[planId].materials[genMatName] = specificId;
      }
    }
  }
  return plans;
}

function mapManufacturing(sheet) {
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  let mfg = {}; 
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    const validFrom = row[0];
    const validTo = row[1];
    if (!validFrom || !validTo || String(validFrom).trim() === "" || String(validTo).trim() === "") {
      continue; 
    }

    const matName = String(row[2]).trim(); 
    const yieldQty = Number(row[4]) || 0;  
    if (!matName || yieldQty <= 0) continue;
    
    mfg[matName] = { yield: yieldQty, rawMats: {} };
    
    for (let col = 5; col < headers.length; col++) {
      let rawMatName = String(headers[col]).trim();
      let qty = Number(row[col]) || 0;
      if (rawMatName && qty > 0) {
        mfg[matName].rawMats[rawMatName] = qty;
      }
    }
  }
  return mfg;
}

// --- PHASE 2: FIFO QUEUE ENGINE (WITH LPP FAILSAFE) ---

function buildFIFOQueues(sheet) {
  // NEW: Added lastPrices object to track the most recent invoice price
  if (!sheet) return { bySpecId: {}, byMatName: {}, lastPrices: { bySpecId: {}, byMatName: {} } };
  const data = sheet.getDataRange().getValues();
  
  let queues = { bySpecId: {}, byMatName: {}, lastPrices: { bySpecId: {}, byMatName: {} } }; 
  let rows = data.slice(1).sort((a, b) => new Date(a[0]) - new Date(b[0]));
  
  rows.forEach(row => {
    let date = row[0];
    let type = String(row[2]).trim(); 
    let matName = String(row[3]).trim(); 
    let specId = String(row[4]).trim(); 
    let qty = Number(row[5]) || 0;
    let price = Number(row[6]) || 0;
    
    if (type === "استلام") {
      let batch = { qty: qty, price: price, date: date }; 
      
      if (specId) {
        if (!queues.bySpecId[specId]) queues.bySpecId[specId] = [];
        queues.bySpecId[specId].push(batch);
        queues.lastPrices.bySpecId[specId] = price; // Constantly updates to the newest price
      }
      if (matName) {
        if (!queues.byMatName[matName]) queues.byMatName[matName] = [];
        queues.byMatName[matName].push(batch);
        queues.lastPrices.byMatName[matName] = price; // Constantly updates to the newest price
      }
    } 
    else if (type.includes("ارتجاع")) {
      let targetQueue = specId ? queues.bySpecId[specId] : queues.byMatName[matName];
      if (targetQueue) {
        let remainingReturn = qty;
        for (let i = targetQueue.length - 1; i >= 0; i--) {
          if (remainingReturn <= 0) break;
          let batch = targetQueue[i];
          if (batch.qty > 0) {
            let deduct = Math.min(batch.qty, remainingReturn);
            batch.qty -= deduct; 
            remainingReturn -= deduct;
          }
        }
      }
    }
  });
  
  return queues;
}

/**
 * CORE MATH: Consume FIFO, fallback to Last Purchase Price if bucket is empty.
 */
function consumeFIFO(queues, searchKey, isGeneral, requiredQty) {
  let cost = 0;
  let remainingNeeded = requiredQty;
  
  let targetQueue = isGeneral ? queues.byMatName[searchKey] : queues.bySpecId[searchKey];
  
  // 1. Drain Standard FIFO Inventory (if it exists)
  if (targetQueue) {
    for (let i = 0; i < targetQueue.length; i++) {
      if (remainingNeeded <= 0) break;
      
      let batch = targetQueue[i];
      if (batch.qty > 0) {
        let consumed = Math.min(batch.qty, remainingNeeded);
        cost += (consumed * batch.price);
        batch.qty -= consumed; 
        remainingNeeded -= consumed;
      }
    }
  }
  
  // 2. FAILSAFE: Bucket is Empty or Negative Inventory
  // If we still need quantity, price the remainder using the Last Known Price
  if (remainingNeeded > 0) {
    let lastPriceMap = isGeneral ? queues.lastPrices.byMatName : queues.lastPrices.bySpecId;
    let lastPrice = lastPriceMap[searchKey] || 0; 
    cost += (remainingNeeded * lastPrice);
  }
  
  return cost;
}

function syncCostLedger() {
  const ss = SpreadsheetApp.openById("1NTLovSrQLtFfebXrSOuWitB29VUV4ifLHmc-Rt_MxWo");
  const costSheet = ss.getSheetByName("Items costs");
  const palletsSheet = ss.getSheetByName("Pallets");

  const maps = buildCostEngineMaps();

  // 1. Map Existing Records & Find Unique Working Days
  const existingData = costSheet.getDataRange().getValues();
  let existingRecords = {};
  let uniqueDates = new Set();

  if (existingData.length > 1) {
    for (let i = 1; i < existingData.length; i++) {
      let rawDate = existingData[i][0];
      
      // Normalize date to midnight so different hours group as the same "Working Day"
      let d = new Date(rawDate);
      d.setHours(0, 0, 0, 0); 
      let dateMs = d.getTime();
      
      let qty = Number(existingData[i][3]) || 0;
      let unitCost = Number(existingData[i][4]) || 0;
      let planId = String(existingData[i][7]).trim(); // Col H
      let unitMarket = Number(existingData[i][8]) || 0;

      if (planId) {
        // FIX: Create an array so multiple pallets with the same PlanID are queued up
        if (!existingRecords[planId]) existingRecords[planId] = [];
        existingRecords[planId].push({ 
          rowIdx: i + 1, dateMs: dateMs, qty: qty, unitCost: unitCost, unitMarket: unitMarket 
        });
        uniqueDates.add(dateMs);
      }
    }
  }

  // FIX: Threshold is now the 2nd most recent unique working day in the ledger
  let sortedDates = Array.from(uniqueDates).sort((a, b) => b - a);
  const thresholdMs = sortedDates.length > 1 ? sortedDates[1] : (sortedDates.length === 1 ? sortedDates[0] : 0);

  // 2. Extract ALL Valid Pallets
  const palletsData = palletsSheet.getDataRange().getValues();
  let allPallets = [];

  for (let i = 1; i < palletsData.length; i++) {
    if (String(palletsData[i][0]).toUpperCase() !== "TRUE") continue; 
    
    let date = palletsData[i][2]; 
    let qty = Number(palletsData[i][4]) || 0; 
    
    // NEW RULE: Completely omit this production run if the quantity is 0
    if (qty <= 0) continue; 
    
    let conciseName = String(palletsData[i][5]).trim(); 
    let planId = String(palletsData[i][6]).trim(); 
    let recipeId = String(palletsData[i][7]).trim(); 

    if (planId) {
      allPallets.push({ date, qty, conciseName, planId, recipeId });
    }
  }

  // CRITICAL: Sort chronologically to strictly respect FIFO
  allPallets.sort((a, b) => new Date(a.date) - new Date(b.date));

  // 3. Execute the Math
  let newRows = [];
  let updateRows = [];

  allPallets.forEach(p => {
    let fullName = maps.products[p.conciseName] || p.conciseName; 
    let pDateMs = new Date(p.date).getTime(); // NEW: Capture pallet production date
    
    let totalActualCost = 0;
    let totalMarketCost = 0;
    let bomActual = {};
    let bomMarket = {};

    let recipe = maps.recipes[p.recipeId];
    let plan = maps.plans[p.planId];

    if (recipe && plan) {
      Object.keys(recipe.materials).forEach(genMat => {
        let reqQtyPerUnit = recipe.materials[genMat];
        let totalReqQty = reqQtyPerUnit * p.qty;

        if (maps.manufacturing[genMat]) { 
          let mfgDetails = maps.manufacturing[genMat];
          let batchesNeeded = totalReqQty / mfgDetails.yield; 
          
          let mfgActual = 0;
          let mfgMarket = 0;
          
          Object.keys(mfgDetails.rawMats).forEach(rawMat => {
             let rawQtyNeeded = mfgDetails.rawMats[rawMat] * batchesNeeded;
             
             let rawActualCost = consumeFIFO(maps.fifoQueues, rawMat, true, rawQtyNeeded);
             mfgActual += rawActualCost;
             bomActual[rawMat] = p.qty > 0 ? Number((rawActualCost / p.qty).toFixed(1)) : 0; 
             
             // Market Math (Match Actual UNLESS Manual Update is active)
             let marketData = getBestMarketPrice(null, maps.marketPrices.byMatName[rawMat], pDateMs);
             let rawMarketCost = 0;
             if (marketData && marketData.source === 'Manual') {
               rawMarketCost = marketData.price * rawQtyNeeded;
             } else {
               rawMarketCost = rawActualCost; // Merge the lines!
             }
             mfgMarket += rawMarketCost;
             bomMarket[rawMat] = p.qty > 0 ? Number((rawMarketCost / p.qty).toFixed(1)) : 0;
          });
          totalActualCost += mfgActual;
          totalMarketCost += mfgMarket;
        }
        else {
          let specId = plan.materials[genMat];
          
          let actualCost = 0;
          let specificQueue = specId ? maps.fifoQueues.bySpecId[specId] : null;
          let hasSpecificInventory = specificQueue && specificQueue.some(batch => batch.qty > 0);
          let hasSpecificPrice = specId ? maps.fifoQueues.lastPrices.bySpecId[specId] : null;

          if (hasSpecificInventory || hasSpecificPrice) { 
            actualCost = consumeFIFO(maps.fifoQueues, specId, false, totalReqQty);
          } else {
            actualCost = consumeFIFO(maps.fifoQueues, genMat, true, totalReqQty);
          }
          totalActualCost += actualCost;
          bomActual[genMat] = p.qty > 0 ? Number((actualCost / p.qty).toFixed(1)) : 0;
          
          // Market Math (Match Actual UNLESS Manual Update is active)
          let marketData = getBestMarketPrice(
             specId ? maps.marketPrices.bySpecId[specId] : null,
             genMat ? maps.marketPrices.byMatName[genMat] : null,
             pDateMs
          );

          let marketCost = 0;
          if (marketData && marketData.source === 'Manual') {
            marketCost = marketData.price * totalReqQty;
          } else {
            marketCost = actualCost; // Merge the lines!
          }
          
          totalMarketCost += marketCost;
          bomMarket[genMat] = p.qty > 0 ? Number((marketCost / p.qty).toFixed(1)) : 0;
        }
      });
    }

    let unitActual = p.qty > 0 ? Number((totalActualCost / p.qty).toFixed(1)) : 0;
    let unitMarket = p.qty > 0 ? Number((totalMarketCost / p.qty).toFixed(1)) : 0;
    
    let rowData = [
      p.date, fullName, "", p.qty, 
      unitActual, totalActualCost, JSON.stringify(bomActual), p.planId,
      unitMarket, totalMarketCost, JSON.stringify(bomMarket)
    ];

    // FIX: Safely pull the FIRST un-updated row for this PlanID, matching 1-to-1
    let existingList = existingRecords[p.planId];
    let existing = existingList && existingList.length > 0 ? existingList.shift() : null;
    
    if (!existing) {
      newRows.push(rowData);
    } else {
      if (existing.dateMs >= thresholdMs) {
        if (existing.qty !== p.qty || existing.unitCost !== unitActual || existing.unitMarket !== unitMarket) {
          updateRows.push({ rowIdx: existing.rowIdx, data: rowData });
        }
      }
    }
  });

  // 4. Write to Sheet
  if (newRows.length > 0) {
    costSheet.getRange(costSheet.getLastRow() + 1, 1, newRows.length, 11).setValues(newRows);
  }
  
  if (updateRows.length > 0) {
    updateRows.forEach(u => {
      costSheet.getRange(u.rowIdx, 1, 1, 11).setValues([u.data]);
    });
  }

  let msg = `Cost Engine Synced. Added ${newRows.length} new pallets.`;
  if (updateRows.length > 0) msg += ` Updated ${updateRows.length} recent pallets.`;
  return msg;
}

// --- PHASE 1.5: MARKET PRICE ENGINE (WITH SOURCE TRACKING) ---

function buildMarketPrices(purchasesSheet, updatesSheet) {
  let marketMap = { bySpecId: {}, byMatName: {} };
  let timeline = [];

  // 1. Extract Actual Purchase Prices
  if (purchasesSheet) {
    const pData = purchasesSheet.getDataRange().getValues();
    for(let i = 1; i < pData.length; i++) {
      let type = String(pData[i][2]).trim();
      if(type === "استلام") {
        timeline.push({
          date: new Date(pData[i][0]).getTime(),
          matName: String(pData[i][3]).trim(),
          specId: String(pData[i][4]).trim(),
          price: Number(pData[i][6]) || 0,
          source: 'Purchase' // TAG: Real Purchase
        });
      }
    }
  }

  // 2. Extract Manual Price Updates
  if (updatesSheet) {
    const uData = updatesSheet.getDataRange().getValues();
    for(let i = 1; i < uData.length; i++) {
      let dateVal = uData[i][0];
      let matName = String(uData[i][1]).trim();
      let price = Number(uData[i][2]) || 0;
      if (dateVal && matName && price > 0) {
        timeline.push({
          date: new Date(dateVal).getTime(),
          matName: matName,
          specId: "", 
          price: price,
          source: 'Manual' // TAG: Manual Update
        });
      }
    }
  }

  // 3. Sort Chronologically
  timeline.sort((a, b) => a.date - b.date);

  // 4. Store the history AND the source tag
  timeline.forEach(event => {
    if (event.specId) {
      if (!marketMap.bySpecId[event.specId]) marketMap.bySpecId[event.specId] = [];
      marketMap.bySpecId[event.specId].push({ price: event.price, dateMs: event.date, source: event.source });
    }
    if (event.matName) {
      if (!marketMap.byMatName[event.matName]) marketMap.byMatName[event.matName] = [];
      marketMap.byMatName[event.matName].push({ price: event.price, dateMs: event.date, source: event.source });
    }
  });

  return marketMap;
}

// --- NEW HELPER: GET PRICE & SOURCE AT SPECIFIC DATE ---
function getBestMarketPrice(specHistory, genHistory, targetDateMs) {
  let bestPrice = 0;
  let bestSource = 'Purchase'; // Default assumption
  let latestMs = -1;

  if (specHistory) {
    for (let i = 0; i < specHistory.length; i++) {
      if (specHistory[i].dateMs <= targetDateMs) {
        if (specHistory[i].dateMs > latestMs) {
          latestMs = specHistory[i].dateMs;
          bestPrice = specHistory[i].price;
          bestSource = specHistory[i].source;
        }
      } else break; 
    }
  }
  if (genHistory) {
    for (let i = 0; i < genHistory.length; i++) {
      if (genHistory[i].dateMs <= targetDateMs) {
        if (genHistory[i].dateMs >= latestMs) { 
          latestMs = genHistory[i].dateMs;
          bestPrice = genHistory[i].price;
          bestSource = genHistory[i].source;
        }
      } else break; 
    }
  }
  return { price: bestPrice, source: bestSource };
}

// --- DASHBOARD API ---

function getDashboardCostData() {
  const ss = SpreadsheetApp.openById("1NTLovSrQLtFfebXrSOuWitB29VUV4ifLHmc-Rt_MxWo");
  const costSheet = ss.getSheetByName("Items costs");
  const data = costSheet.getDataRange().getValues();
  
  if (data.length <= 1) return []; // Return empty if no data
  
  let formattedData = [];
  
  // Start from row 1 (skipping headers)
  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    
    // Safety check: skip completely empty rows
    if (!row[0] || !row[1]) continue; 
    
    formattedData.push({
      date: row[0], // Date
      productName: String(row[1]).trim(),
      qty: Number(row[3]) || 0,
      
      // Actual Costs
      unitActual: Number(row[4]) || 0,
      totalActual: Number(row[5]) || 0,
      bomActual: row[6] ? JSON.parse(row[6]) : {},
      
      planId: String(row[7]).trim(),
      
      // Market Costs
      unitMarket: Number(row[8]) || 0,
      totalMarket: Number(row[9]) || 0,
      bomMarket: row[10] ? JSON.parse(row[10]) : {}
    });
  }
  
  // Sort by date descending (newest first) for the dashboard
  formattedData.sort((a, b) => new Date(b.date) - new Date(a.date));
  
  return formattedData;
}
