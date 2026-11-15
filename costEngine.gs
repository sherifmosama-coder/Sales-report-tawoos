/**
 * ============================================================================
 * ERP COSTING ENGINE - PHASE 1 & 2
 * Two-Tier FIFO (First-In-First-Out) with Nested Bill of Materials (BOM)
 * ============================================================================
 */

// --- PHASE 1: DATA EXTRACTION & MAPPING ---

function normalizeCostDateMs(raw) {
  if (!raw) return NaN;
  if (raw instanceof Date) {
    if (isNaN(raw.getTime())) return NaN;
    const d = new Date(raw.getTime());
    d.setHours(0, 0, 0, 0);
    return d.getTime();
  }
  const str = String(raw).trim();
  if (!str) return NaN;
  if (/^\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}$/.test(str)) {
    const parts = str.split(/[\/\-]/);
    const d1 = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
    if (!isNaN(d1.getTime())) {
      d1.setHours(0, 0, 0, 0);
      return d1.getTime();
    }
  }
  if (/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/.test(str)) {
    const parts = str.split(/[\/\-]/);
    const first = parseInt(parts[0], 10);
    const second = parseInt(parts[1], 10);
    let year = parseInt(parts[2], 10);
    if (year < 100) year += 2000;
    // Support DD/MM/YYYY (and M/D/YYYY when second > 12)
    const day = second > 12 ? second : first;
    const month = (second > 12 ? first : second) - 1;
    const d2 = new Date(year, month, day);
    if (!isNaN(d2.getTime())) {
      d2.setHours(0, 0, 0, 0);
      return d2.getTime();
    }
  }
  const parsed = new Date(str);
  if (isNaN(parsed.getTime())) return NaN;
  parsed.setHours(0, 0, 0, 0);
  return parsed.getTime();
}

function buildCostEngineMaps(existingSS) {
  const ss = existingSS || SpreadsheetApp.openById("1NTLovSrQLtFfebXrSOuWitB29VUV4ifLHmc-Rt_MxWo");
  const prodInfo = mapProducts(ss.getSheetByName("Products"));
  
  return {
    products: prodInfo.names,
    productVatMap: prodInfo.vatRates,
    recipes: mapRecipes(ss.getSheetByName("RECIPE")),
    plans: mapPlans(ss.getSheetByName("PLANS")),
    manufacturing: mapManufacturing(ss.getSheetByName("MANUFACTURING")),
    // Load the Market Prices (single read of Purchased items & Price Updates)
    marketPrices: buildMarketPrices(ss.getSheetByName("Purchased items"), ss.getSheetByName("Price Updates"))
  };
}

function mapProducts(sheet) {
  if (!sheet) return { names: {}, vatRates: {} };
  const data = sheet.getDataRange().getValues();
  let names = {};
  let vatRates = {};
  for (let i = 1; i < data.length; i++) {
    let concise = String(data[i][1]).trim(); 
    let full = String(data[i][5]).trim();    
    let vat = Number(data[i][6]) || 0;
    if (concise) {
      names[concise] = full;
      vatRates[concise] = vat;
    }
  }
  return { names: names, vatRates: vatRates };
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
    
    const product = String(row[2] || "").trim();
    const validFromRaw = row[3];
    const validToRaw = (row[4] !== undefined && row[4] !== null && String(row[4]).trim() !== "") ? row[4] : null;
    const vFromMs = validFromRaw ? normalizeCostDateMs(validFromRaw) : NaN;
    const vToMs = validToRaw ? normalizeCostDateMs(validToRaw) : NaN;

    recipes[recipeId] = {
      product: product,
      validFromMs: isNaN(vFromMs) ? 0 : vFromMs,
      validToMs: isNaN(vToMs) ? Infinity : (vToMs + 86399999),
      materials: {}
    };
    
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

    const product = String(row[3] || "").trim();
    const recipeId = String(row[5] || "").trim();
    
    plans[planId] = { product: product, recipeId: recipeId, materials: {} };
    
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

    const validFromMs = normalizeCostDateMs(validFrom);
    const validToMs = normalizeCostDateMs(validTo);
    const matName = String(row[2]).trim(); 
    const yieldQty = Number(row[4]) || 0;  
    if (!matName || yieldQty <= 0) continue;
    
    const rule = {
      validFromMs: isNaN(validFromMs) ? 0 : validFromMs,
      validToMs: isNaN(validToMs) ? Infinity : (validToMs + 86399999),
      yield: yieldQty,
      rawMats: {}
    };
    
    for (let col = 5; col < headers.length; col++) {
      let rawMatName = String(headers[col]).trim();
      let qty = Number(row[col]) || 0;
      if (rawMatName && qty > 0) {
        rule.rawMats[rawMatName] = qty;
      }
    }

    if (!mfg[matName]) mfg[matName] = [];
    mfg[matName].push(rule);
  }
  return mfg;
}

function resolveManufacturingRule(mfgMap, matName, targetDateMs) {
  const rules = mfgMap ? mfgMap[matName] : null;
  if (!rules || rules.length === 0) return null;
  for (let i = rules.length - 1; i >= 0; i--) {
    if (targetDateMs >= rules[i].validFromMs && targetDateMs <= rules[i].validToMs) {
      return rules[i];
    }
  }
  return rules[rules.length - 1];
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
        let dateMs = normalizeCostDateMs(rawDate);
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
        let dateMs = normalizeCostDateMs(rawDate);
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

function getMaterialRates(targetDateMs, specMap, genMap, hasSpecId) {
  let bestAct = null; let bestMkt = null;

  const findLatestActual = (map) => {
    if (!map || !map.actuals) return null;
    for (let i = map.actuals.length - 1; i >= 0; i--) {
      if (map.actuals[i].dateMs <= targetDateMs) {
        return map.actuals[i];
      }
    }
    return null;
  };

  const findBestUpdate = (map) => {
    if (!map || !map.updates) return;
    for (let i = map.updates.length - 1; i >= 0; i--) {
      if (map.updates[i].dateMs <= targetDateMs) {
        if (!bestMkt || map.updates[i].dateMs > bestMkt.dateMs) bestMkt = map.updates[i];
        break;
      }
    }
  };

  // specId strictly takes precedence over general material when prior actuals exist for specId
  if (hasSpecId) {
    bestAct = findLatestActual(specMap);
    if (!bestAct) bestAct = findLatestActual(genMap);
  } else {
    bestAct = findLatestActual(genMap);
  }

  findBestUpdate(genMap);
  if (hasSpecId) findBestUpdate(specMap);

  const hasActiveOverride = !!(bestMkt && (!bestAct || bestMkt.dateMs > bestAct.dateMs));

  return {
    actBaseRate: bestAct ? bestAct.basePrice : 0,
    actVatRate: bestAct ? bestAct.vatPrice : 0,
    mktRate: hasActiveOverride ? bestMkt.price : 0,
    hasMktOverride: hasActiveOverride,
    unit: bestAct ? bestAct.unit : '',
    actDateMs: bestAct ? bestAct.dateMs : null,
    mktDateMs: hasActiveOverride ? bestMkt.dateMs : null
  };
}

function syncCostLedger(existingSS) {
  const ss = existingSS || SpreadsheetApp.openById("1NTLovSrQLtFfebXrSOuWitB29VUV4ifLHmc-Rt_MxWo");
  const costSheet = ss.getSheetByName("Items costs");
  const palletsSheet = ss.getSheetByName("Pallets");
  const maps = buildCostEngineMaps(ss);

  // 1. EXTRACT PRODUCT VAT RATIOS (reused from Products read in buildCostEngineMaps)
  const productVatMap = maps.productVatMap || {};

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
  let dailyPallets = {};           // dateMs -> { "conciseName||recipeId": { conciseName, planId, recipeId, dateMs } }
  let productRecipeTimeline = {};  // conciseName -> { recipeId: [ { dateMs, planId, recipeId }, ... ] }
  let palletDateToProducts = {};   // dateMs -> Set of conciseNames produced on this day
  let matToProducts = {}; 

  // Strict Date Parser: Only matches authentic Date objects or DD/MM/YYYY / YYYY-MM-DD strings
  var parsePalletDateMs = function(row) {
    for (var col = 0; col <= 4; col++) {
      var raw = row[col];
      if (!raw || raw === true || raw === false) continue;
      
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

  const isSheetError = (val) => /^#(REF!|N\/A|VALUE!|NAME\?|ERROR!|DIV\/0!)/i.test(val);

  for (let i = 1; i < palletsData.length; i++) {
    if (String(palletsData[i][0]).toUpperCase() !== "TRUE") continue; 
    let conciseName = String(palletsData[i][5] || '').trim(); 
    let planId = String(palletsData[i][6] || '').trim(); 
    let recipeId = String(palletsData[i][7] || '').trim(); 
    
    if (!conciseName || !planId || !recipeId) continue;
    if (isSheetError(conciseName) || isSheetError(planId) || isSheetError(recipeId)) continue;

    // Cross-validate that plan and recipe exist and belong to this exact product
    let planObj = maps.plans[planId];
    let recipeObj = maps.recipes[recipeId];
    if (!planObj || !recipeObj) continue;
    if (planObj.product && planObj.product !== conciseName) continue;
    if (recipeObj.product && recipeObj.product !== conciseName) continue;
    if (planObj.recipeId && planObj.recipeId !== recipeId) continue;

    let rowDateMs = parsePalletDateMs(palletsData[i]);
    
    if (!dailyPallets[rowDateMs]) dailyPallets[rowDateMs] = {};
    // "LATEST ON DATE WINS" per (product, recipeId): Downward loop naturally overwrites earlier rows for same product+recipe on the same date
    dailyPallets[rowDateMs][conciseName + "||" + recipeId] = { conciseName: conciseName, planId: planId, recipeId: recipeId, dateMs: rowDateMs };

    // Register pallet production dates into the master calculation calendar
    if (rowDateMs > 0) {
      eventDates.add(rowDateMs);
      if (!palletDateToProducts[rowDateMs]) palletDateToProducts[rowDateMs] = new Set();
      palletDateToProducts[rowDateMs].add(conciseName);
    }
  }

  const baseDatesMs = Array.from(eventDates);
  const minBaseDateMs = baseDatesMs.length > 0 ? Math.min(...baseDatesMs) : 0;
  const maxBaseDateMs = baseDatesMs.length > 0 ? Math.max(...baseDatesMs) : 0;

  // Build sorted chronological timeline per (product, recipeId)
  Object.keys(dailyPallets).forEach(dateKey => {
    let dayEntries = dailyPallets[dateKey];
    Object.keys(dayEntries).forEach(compKey => {
      let entry = dayEntries[compKey];
      if (!productRecipeTimeline[entry.conciseName]) productRecipeTimeline[entry.conciseName] = {};
      if (!productRecipeTimeline[entry.conciseName][entry.recipeId]) productRecipeTimeline[entry.conciseName][entry.recipeId] = [];
      productRecipeTimeline[entry.conciseName][entry.recipeId].push({ planId: entry.planId, recipeId: entry.recipeId, dateMs: entry.dateMs });
    });
  });

  Object.keys(productRecipeTimeline).forEach(conciseName => {
    let byRecipe = productRecipeTimeline[conciseName];
    Object.keys(byRecipe).forEach(rId => {
      byRecipe[rId].sort((a, b) => a.dateMs - b.dateMs);
      
      // Map material dependencies across all recipe versions and manufacturing date ranges for this product
      byRecipe[rId].forEach(cfg => {
        let recipe = maps.recipes[cfg.recipeId];
        let plan = maps.plans[cfg.planId];
        if (recipe) {
          Object.keys(recipe.materials).forEach(genMat => {
            if (!matToProducts[genMat]) matToProducts[genMat] = new Set();
            matToProducts[genMat].add(conciseName);
            
            if (maps.manufacturing[genMat] && maps.manufacturing[genMat].length > 0) {
              maps.manufacturing[genMat].forEach(rule => {
                if (rule.validFromMs > minBaseDateMs && rule.validFromMs <= maxBaseDateMs) {
                  eventDates.add(rule.validFromMs);
                  if (!dateToMats[rule.validFromMs]) dateToMats[rule.validFromMs] = new Set();
                  dateToMats[rule.validFromMs].add(genMat);
                }
                Object.keys(rule.rawMats).forEach(raw => {
                  if (!matToProducts[raw]) matToProducts[raw] = new Set();
                  matToProducts[raw].add(conciseName);
                });
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
  });

  // Helper: Resolves ALL valid active recipe configs for conciseName on targetDateMs (that have appeared in Pallets/PLANS by targetDateMs)
  var getActiveProductConfigs = function(conciseName, targetDateMs) {
    var byRecipe = productRecipeTimeline[conciseName];
    if (!byRecipe) return [];
    var usedConfigs = [];

    Object.keys(byRecipe).forEach(function(rId) {
      var timeline = byRecipe[rId];
      if (!timeline || timeline.length === 0) return;
      if (timeline[0].dateMs > targetDateMs) return;

      var activeCfg = timeline[0];
      for (var k = 0; k < timeline.length; k++) {
        if (timeline[k].dateMs <= targetDateMs) {
          activeCfg = timeline[k];
        } else {
          break;
        }
      }
      usedConfigs.push(activeCfg);
    });

    if (usedConfigs.length === 0) return [];

    var validConfigs = usedConfigs.filter(function(cfg) {
      var rec = maps.recipes[cfg.recipeId];
      if (!rec) return false;
      return targetDateMs >= rec.validFromMs && targetDateMs <= rec.validToMs;
    });

    var result = validConfigs.length > 0 ? validConfigs : usedConfigs;
    // Emit most recently produced recipe LAST so Profitability binary search uses latest production recipe
    result.sort(function(a, b) { return a.dateMs - b.dateMs; });
    return result;
  };

  let sortedDatesMs = Array.from(eventDates).sort((a, b) => a - b);
  if (sortedDatesMs.length === 0) return "No valid calculation dates found.";
  
  const minDateMs = sortedDatesMs[0];
  let allActiveProducts = Object.keys(productRecipeTimeline);

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
        let activeConfigs = getActiveProductConfigs(conciseName, dateMs);
        if (!activeConfigs || activeConfigs.length === 0) return;

        let isVatExempt = (productVatMap[conciseName] === 0);

        activeConfigs.forEach(p => {
          let recipe = maps.recipes[p.recipeId]; 
          let plan = maps.plans[p.planId];
          if (!recipe || !plan) return;

          let totalAct = 0; let totalMkt = 0;
          let bomAct = {}; let bomMkt = {}; 
          let bomMeta = { _isVatExempt: isVatExempt, _recipeId: p.recipeId };
          if (isFinite(recipe.validToMs)) bomMeta._recipeValidToMs = recipe.validToMs;

          const addBomLine = (matKey, qty, r, specId) => {
            let actRateToUse = isVatExempt ? r.actVatRate : r.actBaseRate;
            let mktRateToUse = r.hasMktOverride ? r.mktRate : actRateToUse;
            let lineAct = actRateToUse * qty;
            let lineMkt = mktRateToUse * qty;

            totalAct += lineAct;
            totalMkt += lineMkt;
            bomAct[matKey] = Number(((bomAct[matKey] || 0) + lineAct).toFixed(2));
            bomMkt[matKey] = Number(((bomMkt[matKey] || 0) + lineMkt).toFixed(2));

            if (bomMeta[matKey]) {
              bomMeta[matKey].qty = Number((bomMeta[matKey].qty + qty).toFixed(6));
            } else {
              let metaObj = {
                qty: Number(qty.toFixed(6)),
                actRate: actRateToUse,
                mktRate: mktRateToUse,
                unit: r.unit,
                actDateMs: r.actDateMs,
                mktDateMs: r.mktDateMs,
                hasMktOverride: r.hasMktOverride,
                actBaseRate: r.actBaseRate,
                actVatRate: r.actVatRate
              };
              if (specId) metaObj.specId = specId;
              bomMeta[matKey] = metaObj;
            }
          };

          Object.keys(recipe.materials).forEach(genMat => {
            let reqQty = recipe.materials[genMat];
            let mfg = resolveManufacturingRule(maps.manufacturing, genMat, dateMs);
            if (mfg) {
              Object.keys(mfg.rawMats).forEach(raw => {
                let qty = mfg.rawMats[raw] * (reqQty / mfg.yield);
                let r = getMaterialRates(dateMs, null, maps.marketPrices.byMatName[raw], false);
                addBomLine(raw, qty, r, null);
              });
            } else {
              let specId = plan.materials[genMat];
              let hasSpecId = !!specId;
              let r = getMaterialRates(dateMs, hasSpecId ? maps.marketPrices.bySpecId[specId] : null, maps.marketPrices.byMatName[genMat], hasSpecId);
              addBomLine(genMat, reqQty, r, specId);
            }
          });
          
          if (totalAct > 0) {
            newRows.push([dateStr, maps.products[conciseName] || conciseName, "", 1, Number(totalAct.toFixed(2)), totalAct, JSON.stringify(bomAct), p.planId, Number(totalMkt.toFixed(2)), totalMkt, JSON.stringify(bomMkt), JSON.stringify(bomMeta)]);
          }
        });
     });
  });

  // 5. DEEP CLEAN & WRITE ITEMS COSTS
  costSheet.getRange(2, 1, Math.max(costSheet.getMaxRows()-1, 1), 12).clearContent();
  if (newRows.length > 0) costSheet.getRange(2, 1, newRows.length, 12).setValues(newRows);

  return "Cost Engine Synced (Daily Scan + Active Recipes Resolved).";
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
