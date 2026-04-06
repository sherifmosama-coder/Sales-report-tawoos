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

// HELPER: CALCULATES THE EXACT COST AT A SPECIFIC DATE
function getMaterialRates(targetDateMs, specData, genData) {
  let bestActPrice = 0; let bestActDate = -1; let bestUnit = '';
  let bestUpdPrice = 0; let bestUpdDate = -1;

  // Prioritize Specific ID Matches
  if (specData) {
    for (let i = 0; i < specData.actuals.length; i++) {
      if (specData.actuals[i].dateMs <= targetDateMs) { bestActPrice = specData.actuals[i].price; bestActDate = specData.actuals[i].dateMs; if (specData.actuals[i].unit) bestUnit = specData.actuals[i].unit; } else break;
    }
    for (let i = 0; i < specData.updates.length; i++) {
      if (specData.updates[i].dateMs <= targetDateMs) { bestUpdPrice = specData.updates[i].price; bestUpdDate = specData.updates[i].dateMs; } else break;
    }
  }

  // Fallback to Generic Material Name Matches
  if (genData) {
    if (bestActDate === -1) {
      for (let i = 0; i < genData.actuals.length; i++) {
        if (genData.actuals[i].dateMs <= targetDateMs) { bestActPrice = genData.actuals[i].price; bestActDate = genData.actuals[i].dateMs; if (genData.actuals[i].unit) bestUnit = genData.actuals[i].unit; } else break;
      }
    }
    for (let i = 0; i < genData.updates.length; i++) {
      if (genData.updates[i].dateMs <= targetDateMs) {
        if (genData.updates[i].dateMs >= bestUpdDate) { bestUpdPrice = genData.updates[i].price; bestUpdDate = genData.updates[i].dateMs; }
      } else break;
    }
  }

  // CORE LOGIC: Market mirrors Actual, unless Manual Update is newer!
  let actRate = bestActPrice;
  let mktRate = bestActPrice; 
  if (bestUpdDate >= bestActDate && bestUpdPrice > 0) {
    mktRate = bestUpdPrice;
  }

  return { actRate: actRate, mktRate: mktRate, unit: bestUnit };
}

function buildMarketPrices(purchasesSheet, updatesSheet) {
  let priceMaps = { bySpecId: {}, byMatName: {} };

  // 1. EXTRACT ACTUAL PURCHASES (Strict Validation)
  if (purchasesSheet) {
    const pData = purchasesSheet.getDataRange().getValues();
    for(let i = 1; i < pData.length; i++) {
      let rawDate = pData[i][0];
      let matName = String(pData[i][3] || "").trim();
      let specId = String(pData[i][4] || "").trim();
      let price = Number(pData[i][6]) || 0;
      let unit = String(pData[i][7] || "").trim();

      // STOPS GHOST ROWS: Only process if there is a Date, a Name, and a Price > 0
      if (rawDate && matName.length > 1 && price > 0) {
        let dateMs = new Date(rawDate).getTime();
        if (isNaN(dateMs)) continue;

        if (specId) {
          if (!priceMaps.bySpecId[specId]) priceMaps.bySpecId[specId] = { actuals: [], updates: [] };
          priceMaps.bySpecId[specId].actuals.push({ dateMs: dateMs, price: price, unit: unit });
        }
        if (matName) {
          if (!priceMaps.byMatName[matName]) priceMaps.byMatName[matName] = { actuals: [], updates: [] };
          priceMaps.byMatName[matName].actuals.push({ dateMs: dateMs, price: price, unit: unit });
        }
      }
    }
  }

  // 2. EXTRACT MANUAL OVERRIDES (Strict Validation)
  if (updatesSheet) {
    const uData = updatesSheet.getDataRange().getValues();
    for(let i = 1; i < uData.length; i++) {
      let rawDate = uData[i][0];
      let matName = String(uData[i][1] || "").trim();
      let price = Number(uData[i][2]) || 0;

      if (rawDate && matName.length > 1 && price > 0) {
        let dateMs = new Date(rawDate).getTime();
        if (isNaN(dateMs)) continue;

        if (!priceMaps.byMatName[matName]) priceMaps.byMatName[matName] = { actuals: [], updates: [] };
        priceMaps.byMatName[matName].updates.push({ dateMs: dateMs, price: price });
      }
    }
  }

  Object.values(priceMaps.bySpecId).forEach(m => { m.actuals.sort((a,b)=>a.dateMs - b.dateMs); m.updates.sort((a,b)=>a.dateMs - b.dateMs); });
  Object.values(priceMaps.byMatName).forEach(m => { m.actuals.sort((a,b)=>a.dateMs - b.dateMs); m.updates.sort((a,b)=>a.dateMs - b.dateMs); });

  return priceMaps;
}

function syncCostLedger() {
  const ss = SpreadsheetApp.openById("1NTLovSrQLtFfebXrSOuWitB29VUV4ifLHmc-Rt_MxWo");
  const costSheet = ss.getSheetByName("Items costs");
  const palletsSheet = ss.getSheetByName("Pallets");
  const maps = buildCostEngineMaps();

  // 1. GATHER DATES AND TRACK WHICH MATERIALS CHANGED ON WHICH DATE
  let dateToMats = {}; // { dateMs: Set(matNames) }
  let eventDates = new Set();

  if (maps.marketPrices) {
    const processPriceMap = (priceMap) => {
      Object.keys(priceMap).forEach(id => {
        const mat = priceMap[id];
        mat.actuals.forEach(a => {
          eventDates.add(a.dateMs);
          if (!dateToMats[a.dateMs]) dateToMats[a.dateMs] = new Set();
          dateToMats[a.dateMs].add(id);
        });
        mat.updates.forEach(u => {
          eventDates.add(u.dateMs);
          if (!dateToMats[u.dateMs]) dateToMats[u.dateMs] = new Set();
          dateToMats[u.dateMs].add(id);
        });
      });
    };
    processPriceMap(maps.marketPrices.byMatName || {});
    processPriceMap(maps.marketPrices.bySpecId || {});
  }

  let sortedDatesMs = Array.from(eventDates).sort((a, b) => a - b);
  if (sortedDatesMs.length === 0) return "No valid purchase/update events found.";
  
  const minDateMs = sortedDatesMs[0];

  // 2. MAP PRODUCTS AND RECIPES
  const palletsData = palletsSheet.getDataRange().getValues();
  let activeProducts = {};
  let matToProducts = {}; // { matName: Set(productConciseNames) }

  for (let i = 1; i < palletsData.length; i++) {
    if (String(palletsData[i][0]).toUpperCase() !== "TRUE") continue; 
    let conciseName = String(palletsData[i][5]).trim(); 
    let planId = String(palletsData[i][6]).trim(); 
    let recipeId = String(palletsData[i][7]).trim(); 
    
    if (planId && recipeId && !activeProducts[conciseName]) {
      activeProducts[conciseName] = { planId, recipeId };
      
      // Build the dependency map
      let recipe = maps.recipes[recipeId];
      let plan = maps.plans[planId];
      if (recipe) {
        Object.keys(recipe.materials).forEach(genMat => {
          if (!matToProducts[genMat]) matToProducts[genMat] = new Set();
          matToProducts[genMat].add(conciseName);
          
          // Check for sub-materials if manufactured
          if (maps.manufacturing[genMat]) {
            Object.keys(maps.manufacturing[genMat].rawMats).forEach(raw => {
              if (!matToProducts[raw]) matToProducts[raw] = new Set();
              matToProducts[raw].add(conciseName);
            });
          }
          // Check for specific IDs
          if (plan && plan.materials[genMat]) {
            let specId = plan.materials[genMat];
            if (!matToProducts[specId]) matToProducts[specId] = new Set();
            matToProducts[specId].add(conciseName);
          }
        });
      }
    }
  }

  // 3. GENERATE TARGETED ROWS
  let newRows = [];
  let timeZone = ss.getSpreadsheetTimeZone() || "GMT";

  sortedDatesMs.forEach(dateMs => {
     let dateStr = Utilities.formatDate(new Date(dateMs), timeZone, "yyyy-MM-dd");
     let changedMats = dateToMats[dateMs];
     
     // Determine which products to calculate for this specific date
     let targetProducts = [];
     if (dateMs === minDateMs) {
       // BASELINE: Every product gets a row on the first day
       targetProducts = Object.keys(activeProducts);
     } else {
       // TARGETED: Only products that use the materials changed on this day
       let affected = new Set();
       changedMats.forEach(m => {
         if (matToProducts[m]) matToProducts[m].forEach(p => affected.add(p));
       });
       targetProducts = Array.from(affected);
     }

     targetProducts.forEach(conciseName => {
        let p = activeProducts[conciseName];
        let recipe = maps.recipes[p.recipeId]; let plan = maps.plans[p.planId];
        if (!recipe || !plan) return;

        let totalAct = 0; let totalMkt = 0;
        let bomAct = {}; let bomMkt = {}; let bomMeta = {};

        Object.keys(recipe.materials).forEach(genMat => {
          let reqQty = recipe.materials[genMat];
          if (maps.manufacturing[genMat]) {
            let mfg = maps.manufacturing[genMat];
            Object.keys(mfg.rawMats).forEach(raw => {
              let qty = mfg.rawMats[raw] * (reqQty / mfg.yield);
              let r = getMaterialRates(dateMs, null, maps.marketPrices.byMatName[raw]);
              bomMeta[raw] = { actRate: r.actRate, mktRate: r.mktRate, unit: r.unit };
              totalAct += (r.actRate * qty); totalMkt += (r.mktRate * qty);
              bomAct[raw] = Number((r.actRate * qty).toFixed(2)); bomMkt[raw] = Number((r.mktRate * qty).toFixed(2));
            });
          } else {
            let specId = plan.materials[genMat];
            let r = getMaterialRates(dateMs, specId ? maps.marketPrices.bySpecId[specId] : null, maps.marketPrices.byMatName[genMat]);
            bomMeta[genMat] = { actRate: r.actRate, mktRate: r.mktRate, unit: r.unit };
            totalAct += (r.actRate * reqQty); totalMkt += (r.mktRate * reqQty);
            bomAct[genMat] = Number((r.actRate * reqQty).toFixed(2)); bomMkt[genMat] = Number((r.mktRate * reqQty).toFixed(2));
          }
        });
        
        if (totalAct > 0) {
          newRows.push([dateStr, maps.products[conciseName] || conciseName, "", 1, Number(totalAct.toFixed(2)), totalAct, JSON.stringify(bomAct), p.planId, Number(totalMkt.toFixed(2)), totalMkt, JSON.stringify(bomMkt), JSON.stringify(bomMeta)]);
        }
     });
  });

  // 4. DEEP CLEAN & WRITE
  costSheet.getRange(2, 1, Math.max(costSheet.getMaxRows()-1, 1), 12).clearContent();
  if (newRows.length > 0) costSheet.getRange(2, 1, newRows.length, 12).setValues(newRows);

  return "Cost Engine Synced (Baseline + Targeted Mode).";
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
