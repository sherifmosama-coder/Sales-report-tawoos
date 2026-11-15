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

function buildMarketPrices(purchasesSheet, updatesSheet) {
  let priceMaps = { bySpecId: {}, byMatName: {} };

  // 1. EXTRACT ACTUAL PURCHASES (Now Captures Both Base & VAT-Inclusive Prices)
  if (purchasesSheet) {
    const pData = purchasesSheet.getDataRange().getValues();
    for(let i = 1; i < pData.length; i++) {
      let rawDate = pData[i][0];
      let matName = String(pData[i][3] || "").trim();
      let specId = String(pData[i][4] || "").trim();
      
      let basePrice = Number(pData[i][6]) || 0; // Col G: Base Price
      let vatPrice = Number(pData[i][9]) || basePrice; // Col J: VAT Inclusive Price (falls back to base if empty)
      let unit = String(pData[i][7] || "").trim();

      if (rawDate && matName.length > 1 && basePrice > 0) {
        let dateMs = new Date(rawDate).getTime();
        if (isNaN(dateMs)) continue;

        if (specId) {
          if (!priceMaps.bySpecId[specId]) priceMaps.bySpecId[specId] = { actuals: [], updates: [] };
          priceMaps.bySpecId[specId].actuals.push({ dateMs: dateMs, basePrice: basePrice, vatPrice: vatPrice, unit: unit });
        }
        if (matName) {
          if (!priceMaps.byMatName[matName]) priceMaps.byMatName[matName] = { actuals: [], updates: [] };
          priceMaps.byMatName[matName].actuals.push({ dateMs: dateMs, basePrice: basePrice, vatPrice: vatPrice, unit: unit });
        }
      }
    }
  }

  // 2. EXTRACT MANUAL OVERRIDES
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

function getMaterialRates(targetDateMs, specMap, genMap) {
  let bestAct = null; let bestMkt = null;

  const findBest = (map) => {
    if (!map) return;
    if (map.actuals) {
      for (let i = map.actuals.length - 1; i >= 0; i--) {
        if (map.actuals[i].dateMs <= targetDateMs) {
          if (!bestAct || map.actuals[i].dateMs > bestAct.dateMs) bestAct = map.actuals[i];
          break;
        }
      }
    }
    if (map.updates) {
      for (let i = map.updates.length - 1; i >= 0; i--) {
        if (map.updates[i].dateMs <= targetDateMs) {
          if (!bestMkt || map.updates[i].dateMs > bestMkt.dateMs) bestMkt = map.updates[i];
          break;
        }
      }
    }
  };

  findBest(genMap);
  findBest(specMap);

  return {
    actBaseRate: bestAct ? bestAct.basePrice : 0,
    actVatRate: bestAct ? bestAct.vatPrice : 0,
    mktRate: bestMkt ? bestMkt.price : 0,
    hasMktOverride: !!bestMkt,
    unit: bestAct ? bestAct.unit : '',
    actDateMs: bestAct ? bestAct.dateMs : null,
    mktDateMs: bestMkt ? bestMkt.dateMs : null
  };
}

function syncCostLedger() {
  const ss = SpreadsheetApp.openById("1NTLovSrQLtFfebXrSOuWitB29VUV4ifLHmc-Rt_MxWo");
  const costSheet = ss.getSheetByName("Items costs");
  const palletsSheet = ss.getSheetByName("Pallets");
  const productsSheet = ss.getSheetByName("Products");
  const maps = buildCostEngineMaps();

  // 1. EXTRACT PRODUCT VAT RATIOS
  const prodData = productsSheet.getDataRange().getValues();
  let productVatMap = {};
  for (let i = 1; i < prodData.length; i++) {
     let pName = String(prodData[i][1]).trim(); 
     let pVat = Number(prodData[i][6]) || 0;
     if (pName) productVatMap[pName] = pVat;
  }

  // 2. GATHER MATERIAL PURCHASE & MARKET PRICE DATES
  let dateToMats = {}; 
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

  // 3. MAP PALLETS (STRICT DATE PARSER + DAILY OVERWRITE "LATEST WINS")
  const palletsData = palletsSheet.getDataRange().getValues();
  let dailyPallets = {};           // dateMs -> { conciseName: { planId, recipeId, dateMs } }
  let productTimeline = {};        // conciseName -> [ { dateMs, planId, recipeId }, ... ]
  let palletDateToProducts = {};   // dateMs -> Set of conciseNames produced on this day
  let matToProducts = {}; 

  // Strict Date Parser: Only matches authentic Date objects or DD/MM/YYYY / YYYY-MM-DD strings
  var parsePalletDateMs = function(row) {
    for (var col = 0; col <= 4; col++) {
      var raw = row[col];
      if (!raw) continue;
      
      if (raw instanceof Date) {
        if (!isNaN(raw.getTime())) {
          var d = new Date(raw.getTime());
          d.setHours(0, 0, 0, 0);
          return d.getTime();
        }
      }
      
      var str = String(raw).trim();
      // Match YYYY-MM-DD or YYYY/MM/DD
      if (/^\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}$/.test(str)) {
        var parts = str.split(/[\/\-]/);
        var d1 = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
        if (!isNaN(d1.getTime())) {
          d1.setHours(0, 0, 0, 0);
          return d1.getTime();
        }
      }
      // Match DD-MM-YYYY or DD/MM/YYYY
      if (/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/.test(str)) {
        var parts = str.split(/[\/\-]/);
        var day = parseInt(parts[0], 10);
        var month = parseInt(parts[1], 10) - 1;
        var year = parseInt(parts[2], 10);
        if (year < 100) year += 2000;
        var d2 = new Date(year, month, day);
        if (!isNaN(d2.getTime())) {
          d2.setHours(0, 0, 0, 0);
          return d2.getTime();
        }
      }
    }
    return 0; // Fallback for pallet rows without dates
  };

  for (let i = 1; i < palletsData.length; i++) {
    if (String(palletsData[i][0]).toUpperCase() !== "TRUE") continue; 
    let conciseName = String(palletsData[i][5] || '').trim(); 
    let planId = String(palletsData[i][6] || '').trim(); 
    let recipeId = String(palletsData[i][7] || '').trim(); 
    
    if (conciseName && planId && recipeId) {
      let rowDateMs = parsePalletDateMs(palletsData[i]);
      
      if (!dailyPallets[rowDateMs]) dailyPallets[rowDateMs] = {};
      // "LATEST ON DATE WINS": Downward loop naturally overwrites earlier rows on the same date
      dailyPallets[rowDateMs][conciseName] = { planId: planId, recipeId: recipeId, dateMs: rowDateMs };

      // Register pallet production dates into the master calculation calendar
      if (rowDateMs > 0) {
        eventDates.add(rowDateMs);
        if (!palletDateToProducts[rowDateMs]) palletDateToProducts[rowDateMs] = new Set();
        palletDateToProducts[rowDateMs].add(conciseName);
      }
    }
  }

  // Build sorted chronological timeline per product
  Object.keys(dailyPallets).forEach(dateKey => {
    let dMs = Number(dateKey);
    let dayProds = dailyPallets[dateKey];
    Object.keys(dayProds).forEach(conciseName => {
      if (!productTimeline[conciseName]) productTimeline[conciseName] = [];
      productTimeline[conciseName].push(dayProds[conciseName]);
    });
  });

  Object.keys(productTimeline).forEach(conciseName => {
    productTimeline[conciseName].sort((a, b) => a.dateMs - b.dateMs);
    
    // Map material dependencies across all recipe versions for this product
    productTimeline[conciseName].forEach(cfg => {
      let recipe = maps.recipes[cfg.recipeId];
      let plan = maps.plans[cfg.planId];
      if (recipe) {
        Object.keys(recipe.materials).forEach(genMat => {
          if (!matToProducts[genMat]) matToProducts[genMat] = new Set();
          matToProducts[genMat].add(conciseName);
          
          if (maps.manufacturing[genMat]) {
            Object.keys(maps.manufacturing[genMat].rawMats).forEach(raw => {
              if (!matToProducts[raw]) matToProducts[raw] = new Set();
              matToProducts[raw].add(conciseName);
            });
          }
          if (plan && plan.materials[genMat]) {
            let specId = plan.materials[genMat];
            if (!matToProducts[specId]) matToProducts[specId] = new Set();
            matToProducts[specId].add(conciseName);
          }
        });
      }
    });
  });

  // Helper: Resolves active recipe active on or before targetDateMs (latest deployed takes effect)
  var getActiveProductConfig = function(conciseName, targetDateMs) {
    var timeline = productTimeline[conciseName];
    if (!timeline || timeline.length === 0) return null;
    
    var activeCfg = timeline[0];
    for (var k = 0; k < timeline.length; k++) {
      if (timeline[k].dateMs <= targetDateMs) {
        activeCfg = timeline[k];
      } else {
        break;
      }
    }
    return activeCfg;
  };

  let sortedDatesMs = Array.from(eventDates).sort((a, b) => a - b);
  if (sortedDatesMs.length === 0) return "No valid calculation dates found.";
  
  const minDateMs = sortedDatesMs[0];
  let allActiveProducts = Object.keys(productTimeline);

  // 4. GENERATE COST ROWS WITH RECIPE EVOLUTION & INTELLIGENT VAT
  let newRows = [];
  let timeZone = ss.getSpreadsheetTimeZone() || "GMT";

  sortedDatesMs.forEach(dateMs => {
     let dateStr = Utilities.formatDate(new Date(dateMs), timeZone, "yyyy-MM-dd");
     let changedMats = dateToMats[dateMs];
     
     let targetProducts = [];
     if (dateMs === minDateMs) {
       targetProducts = allActiveProducts;
     } else {
       let affected = new Set();
       if (changedMats) {
         changedMats.forEach(m => {
           if (matToProducts[m]) matToProducts[m].forEach(p => affected.add(p));
         });
       }
       if (palletDateToProducts[dateMs]) {
         palletDateToProducts[dateMs].forEach(p => affected.add(p));
       }
       targetProducts = Array.from(affected);
     }

     targetProducts.forEach(conciseName => {
        let p = getActiveProductConfig(conciseName, dateMs);
        if (!p) return;
        let recipe = maps.recipes[p.recipeId]; 
        let plan = maps.plans[p.planId];
        if (!recipe || !plan) return;

        let isVatExempt = (productVatMap[conciseName] === 0);

        let totalAct = 0; let totalMkt = 0;
        let bomAct = {}; let bomMkt = {}; 
        let bomMeta = { _isVatExempt: isVatExempt };

        Object.keys(recipe.materials).forEach(genMat => {
          let reqQty = recipe.materials[genMat];
          if (maps.manufacturing[genMat]) {
            let mfg = maps.manufacturing[genMat];
            Object.keys(mfg.rawMats).forEach(raw => {
              let qty = mfg.rawMats[raw] * (reqQty / mfg.yield);
              let r = getMaterialRates(dateMs, null, maps.marketPrices.byMatName[raw]);
              
              let actRateToUse = isVatExempt ? r.actVatRate : r.actBaseRate;
              let mktRateToUse = r.hasMktOverride ? r.mktRate : actRateToUse;

              bomMeta[raw] = { actRate: actRateToUse, mktRate: mktRateToUse, unit: r.unit, actDateMs: r.actDateMs, mktDateMs: r.mktDateMs, hasMktOverride: r.hasMktOverride, actBaseRate: r.actBaseRate, actVatRate: r.actVatRate };
              totalAct += (actRateToUse * qty); totalMkt += (mktRateToUse * qty);
              bomAct[raw] = Number((actRateToUse * qty).toFixed(2)); bomMkt[raw] = Number((mktRateToUse * qty).toFixed(2));
            });
          } else {
            let specId = plan.materials[genMat];
            let r = getMaterialRates(dateMs, specId ? maps.marketPrices.bySpecId[specId] : null, maps.marketPrices.byMatName[genMat]);
            
            let actRateToUse = isVatExempt ? r.actVatRate : r.actBaseRate;
            let mktRateToUse = r.hasMktOverride ? r.mktRate : actRateToUse;

            bomMeta[genMat] = { actRate: actRateToUse, mktRate: mktRateToUse, unit: r.unit, actDateMs: r.actDateMs, mktDateMs: r.mktDateMs, hasMktOverride: r.hasMktOverride, actBaseRate: r.actBaseRate, actVatRate: r.actVatRate };
            totalAct += (actRateToUse * reqQty); totalMkt += (mktRateToUse * reqQty);
            bomAct[genMat] = Number((actRateToUse * reqQty).toFixed(2)); bomMkt[genMat] = Number((mktRateToUse * reqQty).toFixed(2));
          }
        });
        
        if (totalAct > 0) {
          newRows.push([dateStr, maps.products[conciseName] || conciseName, "", 1, Number(totalAct.toFixed(2)), totalAct, JSON.stringify(bomAct), p.planId, Number(totalMkt.toFixed(2)), totalMkt, JSON.stringify(bomMkt), JSON.stringify(bomMeta)]);
        }
     });
  });

  // 5. DEEP CLEAN & WRITE ITEMS COSTS
  costSheet.getRange(2, 1, Math.max(costSheet.getMaxRows()-1, 1), 12).clearContent();
  if (newRows.length > 0) costSheet.getRange(2, 1, newRows.length, 12).setValues(newRows);

  return "Cost Engine Synced (Daily Scan + Latest Recipe Resolved).";
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

// ============================================================================
// COMMERCIAL TERMS & NET ASP ENGINE (For Production Costing)
// ============================================================================

// ============================================================================
// COMMERCIAL TERMS & NET ASP ENGINE (For Production Costing)
// ============================================================================

function getNetSalesDataFromCache() {
  try {
    let cache = readCacheFile("Tawoos_Cache_NetSales.json");
    // If cache is completely empty or missing, force a rebuild
    if (!cache || !cache.data || cache.data.length === 0) return generateNetSalesCache(true);
    return cache;
  } catch (e) {
    return generateNetSalesCache(true);
  }
}

function generateNetSalesCache(forceFullRebuild = true) {
  forceFullRebuild = true; // Hardcoded to rebuild completely until stable

  const CUTOFF_MS = new Date().getTime() - (12 * 24 * 60 * 60 * 1000); 
  
  const ssData = SpreadsheetApp.getActiveSpreadsheet(); 
  const ssExternal = SpreadsheetApp.openById("1NTLovSrQLtFfebXrSOuWitB29VUV4ifLHmc-Rt_MxWo"); 

  // --- AGGRESSIVE DATA SCRUBBERS ---
  const parseNum = (val) => {
      if (!val) return 0;
      let clean = String(val).replace(/,/g, '').trim(); 
      let num = parseFloat(clean);
      return isNaN(num) ? 0 : num;
  };

  const parseDateMs = (val) => {
      if (!val) return NaN;
      if (val instanceof Date) return val.getTime();
      let ms = new Date(val).getTime();
      if (!isNaN(ms)) return ms;
      let parts = String(val).split(/[\/\-]/);
      if (parts.length === 3) return new Date(`${parts[2]}-${parts[1]}-${parts[0]}`).getTime(); 
      return NaN;
  };

  // 1. EXTRACT CONTRACTS
  let termsSheet = ssExternal.getSheetByName("Commercial Terms");
  let contracts = [];
  if (termsSheet) {
      let tData = termsSheet.getDataRange().getValues();
      for(let i = 1; i < tData.length; i++) {
          let client = String(tData[i][0]).trim();
          let prod = String(tData[i][1]).trim();
          let margin = parseNum(tData[i][2]);
          
          let vFrom = tData[i][3] ? parseDateMs(tData[i][3]) : 0;
          let vTo = tData[i][4] ? parseDateMs(tData[i][4]) : Infinity;

          if(client && margin > 0) {
              let marginDec = margin > 1 ? margin / 100 : margin;
              contracts.push({ client: client, product: prod, margin: marginDec, validFrom: vFrom, validTo: vTo });
          }
      }
  }

  // 2. MAP FULL PRODUCT NAMES TO CONCISE NAMES
  let prodSheet = ssExternal.getSheetByName("Products");
  let fullToConcise = {};
  if(prodSheet) {
      let pData = prodSheet.getDataRange().getValues();
      for(let i = 1; i < pData.length; i++) {
          let concise = String(pData[i][1]).trim();
          let full = String(pData[i][5]).trim();
          if(concise && full) fullToConcise[full] = concise;
      }
  }

  // 3. PARSE "ALL DATA" (THE MASTER LEDGER)
  let allDataSheet = ssData.getSheetByName("All Data");
  if(!allDataSheet) return { error: "All Data sheet not found" };

  let existingCache = forceFullRebuild ? null : readCacheFile("Tawoos_Cache_NetSales.json");
  let output = (existingCache && existingCache.data) ? existingCache.data.filter(item => item.dateMs < CUTOFF_MS) : [];

  let allData = allDataSheet.getDataRange().getValues();
  
  const ROW_PROD_NAMES = 7;     
  const ROW_DATA_START = 9;  
  const COL_BRANCH = 2;         // NEW: Capture Branch
  const COL_MAIN_CLIENT = 3;    
  const COL_SEGMENT = 4;        // NEW: Capture Segment
  const COL_DATE = 6;           
  const COL_PROD_START = 17;    

  if (allData.length <= ROW_PROD_NAMES) return { error: "All Data sheet missing rows" };
  let headers = allData[ROW_PROD_NAMES]; 
  
  let targetCols = [];
  for(let c = COL_PROD_START; c < headers.length; c += 2) { 
      let fullProd = String(headers[c]).trim();
      if(fullProd && fullToConcise[fullProd]) {
          // CRITICAL FIX: We are storing BOTH the concise and the full name here
          targetCols.push({ qtyCol: c, revCol: c + 1, concise: fullToConcise[fullProd], full: fullProd });
      }
  }

  for(let i = ROW_DATA_START; i < allData.length; i++) {
      let branch = String(allData[i][COL_BRANCH]).trim();
      let client = String(allData[i][COL_MAIN_CLIENT]).trim(); 
      let segment = String(allData[i][COL_SEGMENT]).trim();
      let rawDate = allData[i][COL_DATE];               
      
      if(!rawDate || !client) continue;

      let dateMs = parseDateMs(rawDate);
      if (isNaN(dateMs)) continue;
      
      if (existingCache && existingCache.data && dateMs < CUTOFF_MS) continue;

      let dateStr = (rawDate instanceof Date) ?
            rawDate.getFullYear() + "-" + String(rawDate.getMonth() + 1).padStart(2, '0') + "-" + String(rawDate.getDate()).padStart(2, '0') :
            String(rawDate).substring(0, 10);

      targetCols.forEach(tc => {
          let qty = parseNum(allData[i][tc.qtyCol]);
          let grossRev = parseNum(allData[i][tc.revCol]);

          // CRITICAL BUG FIX: Completely ignore zero-price transactions (e.g. free samples/promos)
          // so they do not artificially inflate the divisor and lower the true ASP.
          if(qty > 0 && grossRev > 0) {
              let applicableMargin = 0;
              
              // We use the Concise name to check your Commercial Terms sheet
              let specific = contracts.find(c => c.client === client && c.product === tc.concise && dateMs >= c.validFrom && dateMs <= c.validTo);
              if(specific) {
                  applicableMargin = specific.margin;
              } else {
                  let blanket = contracts.find(c => c.client === client && c.product === "All Products" && dateMs >= c.validFrom && dateMs <= c.validTo);
                  if(blanket) applicableMargin = blanket.margin;
              }

              let netRev = grossRev * (1 - applicableMargin);
              let netAsp = grossRev > 0 ? (netRev / qty) : 0;
              let grossAsp = grossRev > 0 ? (grossRev / qty) : 0;

              output.push({
                  dateMs: dateMs,
                  date: dateStr,
                  client: client,
                  branch: branch || client,
                  segment: segment || 'Unassigned',
                  product: tc.full, 
                  productConcise: tc.concise,
                  qty: qty,
                  grossRev: grossRev,
                  netRev: netRev,
                  grossAsp: grossAsp,
                  netAsp: netAsp,
                  backMarginApplied: applicableMargin
              });
          }
      });
  }

  output.sort((a, b) => b.dateMs - a.dateMs);

  writeCacheFile("Tawoos_Cache_NetSales.json", output);
  
  // Clear the Profitability Cache so it forces a fresh rebuild upon next load
  writeCacheFile("Tawoos_Cache_ProfitabilityMerged.json", {});
  
  return { lastUpdated: new Date().getTime(), latestData: output[0] ? output[0].date : "N/A", data: output };
}

function DIAGNOSE_NET_SALES() {
  const ssData = SpreadsheetApp.getActiveSpreadsheet(); 
  const ssExternal = SpreadsheetApp.openById("1NTLovSrQLtFfebXrSOuWitB29VUV4ifLHmc-Rt_MxWo"); 
  
  Logger.log("--- STARTING DIAGNOSTIC ---");

  // 1. Check the Products translation sheet
  let prodSheet = ssExternal.getSheetByName("Products");
  let pData = prodSheet.getDataRange().getValues();
  let sampleFull = String(pData[1][5]).trim(); // Row 2, Column F
  Logger.log("1. Sample Full Product Name from 'Products' (Col F): [" + sampleFull + "]");
  
  // 2. Check the Master Ledger Headers
  let allDataSheet = ssData.getSheetByName("All Data");
  let allData = allDataSheet.getDataRange().getValues();
  
  let headers = allData[7]; // Index 7 = Row 8
  Logger.log("2. Checking Row 8 in 'All Data' for Headers...");
  Logger.log("   -> Header in Column R (Index 17): [" + String(headers[17]).trim() + "]");
  Logger.log("   -> Header in Column T (Index 19): [" + String(headers[19]).trim() + "]");
  
  if (sampleFull !== String(headers[17]).trim()) {
      Logger.log("   *** WARNING: The Product Name in Col F of 'Products' DOES NOT EXACTLY MATCH the Header in 'All Data' ***");
  }

  // 3. Check the Master Ledger Data
  Logger.log("3. Checking Row 10 in 'All Data' for First Transaction...");
  Logger.log("   -> Client (Col D): [" + String(allData[9][3]).trim() + "]");
  Logger.log("   -> Date (Col G): [" + String(allData[9][6]) + "]");
  Logger.log("   -> Quantity (Col R): [" + String(allData[9][17]) + "]");
  Logger.log("   -> Revenue (Col S): [" + String(allData[9][18]) + "]");
  
  Logger.log("--- DIAGNOSTIC COMPLETE ---");
}

// ==========================================
// PROFITABILITY & INVOICE INSPECTOR ENGINE
// ==========================================
function getProfitabilityMergedData() {
   // 1. ATTEMPT TO LOAD FROM CACHE FIRST (Lightning Fast)
   let cache = readCacheFile("Tawoos_Cache_ProfitabilityMerged.json");
   if (cache && cache.data && cache.data.length > 0) {
       return cache;
   }
   // 2. IF NO CACHE EXISTS, BUILD IT
   return buildProfitabilityCache();
}

// Separate Builder Function to handle the heavy lifting
function buildProfitabilityCache() {
   let salesCache = readCacheFile("Tawoos_Cache_NetSales.json");
   let costCache = readCacheFile("Tawoos_Cache_Costing.json");
   
   if (!salesCache || !salesCache.data) return { error: "Net Sales cache missing. Please sync Net Sales." };
   if (!costCache || !costCache.data) return { error: "Costing cache missing. Please sync Costing." };

   let salesData = salesCache.data;
   let costData = costCache.data;
   let pMap = buildProductHashMap(); 
   
   // 1. Build Chronological Cost Timeline Map
   let costTimeline = {};
   costData.forEach(c => {
      if (!costTimeline[c.product]) costTimeline[c.product] = [];
      costTimeline[c.product].push({ dateMs: new Date(c.date).getTime(), cost: c.unitActual || 0 });
   });
   
   Object.keys(costTimeline).forEach(p => costTimeline[p].sort((a,b) => a.dateMs - b.dateMs));
   
   // 🧠 ENGINE UPGRADE: Binary Search Algorithm
   // This instantly drops the execution time from minutes to milliseconds
   const getCostAtDate = (product, dateMs) => {
      let timeline = costTimeline[product];
      if (!timeline || timeline.length === 0) return 0;
      
      let low = 0;
      let high = timeline.length - 1;
      let activeCost = timeline[0].cost;
      
      while (low <= high) {
          let mid = Math.floor((low + high) / 2);
          if (timeline[mid].dateMs <= dateMs) {
              activeCost = timeline[mid].cost;
              low = mid + 1; // It's valid, but try to find a closer date ahead
          } else {
              high = mid - 1; // Too far ahead, step back
          }
      }
      return activeCost;
   };

   // 2. FETCH ADVANCED COMMERCIAL CONTRACTS
   let contractsData = [];
   try {
       let contractsSS = SpreadsheetApp.openById("11WPuCwQn5LKceSB0NRsk9znEqYc02smdoGt7bLqTY84");
       let contractsSheet = contractsSS.getSheetByName("Contracts");
       if (contractsSheet) {
           let cVals = contractsSheet.getDataRange().getValues();
           for(let i = 1; i < cVals.length; i++) {
               let client = String(cVals[i][0]).trim();
               if (!client) continue;
               
               let isPrivate = cVals[i][4] === true || String(cVals[i][4]).toUpperCase() === 'TRUE';
               let isTawoos = cVals[i][5] === true || String(cVals[i][5]).toUpperCase() === 'TRUE';
               let cDisc = parseFloat(cVals[i][6]) || 0;
               if (cDisc > 1) cDisc = cDisc / 100;
               let iDisc = parseFloat(cVals[i][7]) || 0;
               if (iDisc > 1) iDisc = iDisc / 100;
               let cashAmt = parseFloat(cVals[i][8]) || 0;

               contractsData.push({
                   client: client, isPrivate: isPrivate, isTawoos: isTawoos,
                   contractDisc: cDisc, invoiceDisc: iDisc, cash: cashAmt
               });
           }
       }
   } catch(e) {
       Logger.log("Failed to load Contracts spreadsheet.");
   }

   // 3. MERGE BASE DATA 
   let merged = [];
   salesData.forEach(s => {
      let pMeta = pMap[s.productConcise] || pMap[s.product] || { type: 'OWN PRODUCT', category: 'Uncategorized' };
      let unitCost = getCostAtDate(s.product, s.dateMs);
      
      merged.push({
         dateMs: s.dateMs, date: s.date, client: s.client,
         branch: s.branch, segment: s.segment, category: pMeta.category,
         line: pMeta.type, product: s.product, qty: s.qty,
         grossRev: s.grossRev, unitCost: unitCost
      });
   });
   
   merged.sort((a, b) => b.dateMs - a.dateMs);
   
   let finalOutput = { data: merged, contracts: contractsData, lastUpdated: new Date().getTime() };
   
   // 🧠 CACHE THE RESULTS: We save the finished math to a file, so the next load is instant.
   writeCacheFile("Tawoos_Cache_ProfitabilityMerged.json", finalOutput);
   
   return finalOutput;
}

function generateProfitabilityExportUrl(payloadStr) {
   const tree = JSON.parse(payloadStr);
   const ss = SpreadsheetApp.create("[TEMP EXPORT] Profitability Inspector");
   const sheet = ss.getSheets()[0];
   
   const headers = ["Hierarchy", "Volume (Qty)", "Net Revenue (EGP)", "COGS (EGP)", "Gross Profit (EGP)", "Gross Margin %", "Fixed Fees & OPEX (EGP)", "Net Profit (EGP)", "Net Margin %"];
   sheet.appendRow(headers);
   sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#0284c7").setFontColor("white");
   
   let rowIndex = 2;

   function writeNode(nodeList, depth) {
       nodeList.forEach(node => {
           let indent = Array(depth).fill("      ").join("");
           let gMargin = node.netRev > 0 ? (node.grossProfit / node.netRev * 100).toFixed(1) + "%" : "0%";
           let nMargin = node.netRev > 0 ? (node.netProfit / node.netRev * 100).toFixed(1) + "%" : "0%";
           
           let prefix = depth === 0 ? "" : (node.level === 'product' ? "        • " : "▸ ");
           sheet.appendRow([indent + prefix + node.name, node.qty, node.netRev, node.cogs, node.grossProfit, gMargin, node.opexFees, node.netProfit, nMargin]);
           
           let currentRowIdx = rowIndex;
           let bgColors = ["#f1f5f9", "#f8fafc", "#ffffff", "#ffffff"];
           let fontColors = ["#0f172a", "#334155", "#475569", "#64748b"];
           
           if (node.level === 'product') {
               sheet.getRange(currentRowIdx, 1, 1, headers.length).setFontColor("#64748b");
           } else {
               sheet.getRange(currentRowIdx, 1, 1, headers.length)
                    .setBackground(bgColors[Math.min(depth, bgColors.length-1)])
                    .setFontColor(fontColors[Math.min(depth, fontColors.length-1)])
                    .setFontWeight("bold");
           }
           
           rowIndex++;
           
           if (node.children && Object.keys(node.children).length > 0) {
               let childrenArr = Object.values(node.children).sort((a,b) => b.netRev - a.netRev);
               let childCount = getDescendantCount(childrenArr);
               writeNode(childrenArr, depth + 1);
               if (childCount > 0) {
                   sheet.getRange(currentRowIdx + 1, 1, childCount, 1).shiftRowGroupDepth(depth + 1);
               }
           }
       });
   }

   function getDescendantCount(nodeList) {
       let count = 0;
       nodeList.forEach(n => {
           count++;
           if (n.children && Object.keys(n.children).length > 0) {
               count += getDescendantCount(Object.values(n.children));
           }
       });
       return count;
   }

   let rootNodes = Object.values(tree.children).sort((a,b) => b.netRev - a.netRev);
   writeNode(rootNodes, 0);

   // Add Grand Total Row at the bottom
   let rootGM = tree.netRev > 0 ? (tree.grossProfit / tree.netRev * 100).toFixed(1) + "%" : "0%";
   let rootNM = tree.netRev > 0 ? (tree.netProfit / tree.netRev * 100).toFixed(1) + "%" : "0%";
   sheet.appendRow(["GRAND TOTAL (INCL. GLOBAL OPEX)", tree.qty, tree.netRev, tree.cogs, tree.grossProfit, rootGM, tree.opexFees, tree.netProfit, rootNM]);
   sheet.getRange(rowIndex, 1, 1, headers.length).setFontWeight("bold").setBackground("#0f172a").setFontColor("white");

   sheet.collapseAllRowGroups();
   sheet.autoResizeColumns(1, headers.length);
   if (sheet.getColumnWidth(1) < 300) sheet.setColumnWidth(1, 400);
   
   let url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=xlsx&portrait=false&fitw=true`;
   return { url: url, fileId: ss.getId() };
}
