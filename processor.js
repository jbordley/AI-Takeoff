/**
 * AI Takeoff - Rules Engine Processor
 * 
 * Takes a list of zones (room type + sqft) and produces:
 * 1. Per-zone equipment suggestions
 * 2. Infrastructure rollup (amps, switches, NVR, etc.)
 * 3. Full BOQ with pricing
 */

const fs = require('fs');
const path = require('path');

// Load data
const symbolLibrary = JSON.parse(fs.readFileSync(path.join(__dirname, '../data/symbol-library.json'), 'utf8'));
const rulesEngine = JSON.parse(fs.readFileSync(path.join(__dirname, '../data/rules-engine.json'), 'utf8'));

// Build lookup map
const symbolMap = {};
symbolLibrary.symbols.forEach(s => { symbolMap[s.id] = s; });

/**
 * Evaluate a formula string with sqft variable
 */
function evalFormula(formula, sqft) {
  const safeFormula = formula
    .replace(/sqft/g, sqft)
    .replace(/ceil/g, 'Math.ceil')
    .replace(/max/g, 'Math.max')
    .replace(/min/g, 'Math.min')
    .replace(/floor/g, 'Math.floor');
  try {
    return Math.max(0, Math.round(eval(safeFormula)));
  } catch (e) {
    console.error(`Formula error: ${formula}`, e);
    return 0;
  }
}

/**
 * Process a single zone and return equipment list
 */
function processZone(zone) {
  const { name, type, sqft } = zone;
  const roomRules = rulesEngine.roomTypes[type];
  
  if (!roomRules) {
    console.warn(`Unknown room type: ${type}`);
    return { zone: name, type, sqft, items: [], error: `Unknown room type: ${type}` };
  }

  const items = roomRules.equipment.map(rule => {
    const qty = evalFormula(rule.formula, sqft);
    const symbol = symbolMap[rule.symbolId];
    
    if (!symbol || qty === 0) return null;

    const equipmentCost = qty * symbol.unitCost;
    const laborCost = qty * symbol.laborHours * symbol.laborRate;

    return {
      symbolId: rule.symbolId,
      name: symbol.name,
      partNumber: symbol.partNumber,
      manufacturer: symbol.manufacturer,
      model: symbol.model,
      category: symbol.category,
      qty,
      unitCost: symbol.unitCost,
      equipmentCost,
      laborHours: symbol.laborHours * qty,
      laborRate: symbol.laborRate,
      laborCost,
      totalCost: equipmentCost + laborCost,
      priority: rule.priority,
      description: rule.description,
      mounting: symbol.mounting
    };
  }).filter(Boolean);

  const zoneTotals = items.reduce((acc, item) => ({
    equipment: acc.equipment + item.equipmentCost,
    labor: acc.labor + item.laborCost,
    total: acc.total + item.totalCost,
    laborHours: acc.laborHours + item.laborHours
  }), { equipment: 0, labor: 0, total: 0, laborHours: 0 });

  return { zone: name, type, label: roomRules.label, sqft, items, totals: zoneTotals };
}

/**
 * Calculate infrastructure backbone from total device counts
 */
function calculateInfrastructure(zoneResults) {
  // Count totals across all zones
  let totalSpeakers = 0;
  let totalZones = zoneResults.length;
  let totalPoeDevices = 0;
  let totalCameras = 0;
  let totalReaders = 0;
  let totalDisplays = 0;

  zoneResults.forEach(z => {
    z.items.forEach(item => {
      const sym = symbolMap[item.symbolId];
      if (!sym) return;
      
      if (sym.subcategory === 'Audio' && sym.mounting !== 'rack') totalSpeakers += item.qty;
      if (sym.subcategory === 'Surveillance' && sym.mounting !== 'rack') totalCameras += item.qty;
      if (sym.subcategory === 'Access' && sym.id === 'AC-READER-01') totalReaders += item.qty;
      if (sym.subcategory === 'Video' && sym.mounting === 'wall') totalDisplays += item.qty;
      if (sym.specs && sym.specs.poe) totalPoeDevices += item.qty;
    });
  });

  // Also count APs as PoE
  zoneResults.forEach(z => {
    z.items.forEach(item => {
      if (item.symbolId === 'NET-AP-01') totalPoeDevices += item.qty;
    });
  });

  const infraRules = rulesEngine.infrastructureRules;
  const infraItems = [];

  // Amplifiers
  if (totalSpeakers > 0) {
    const qty = evalFormula(infraRules.amplifiers.formula.replace(/totalSpeakers/g, totalSpeakers), 0);
    const sym = symbolMap[infraRules.amplifiers.symbolId];
    if (sym && qty > 0) {
      infraItems.push({
        symbolId: sym.id, name: sym.name, partNumber: sym.partNumber,
        manufacturer: sym.manufacturer, model: sym.model, category: sym.category,
        qty, unitCost: sym.unitCost, equipmentCost: qty * sym.unitCost,
        laborHours: sym.laborHours * qty, laborRate: sym.laborRate,
        laborCost: qty * sym.laborHours * sym.laborRate,
        totalCost: qty * (sym.unitCost + sym.laborHours * sym.laborRate),
        priority: 'required', description: infraRules.amplifiers.description, mounting: sym.mounting
      });
    }
  }

  // DSP
  if (totalZones > 0) {
    const qty = evalFormula(infraRules.dsp.formula.replace(/totalZones/g, totalZones), 0);
    const sym = symbolMap[infraRules.dsp.symbolId];
    if (sym && qty > 0) {
      infraItems.push({
        symbolId: sym.id, name: sym.name, partNumber: sym.partNumber,
        manufacturer: sym.manufacturer, model: sym.model, category: sym.category,
        qty, unitCost: sym.unitCost, equipmentCost: qty * sym.unitCost,
        laborHours: sym.laborHours * qty, laborRate: sym.laborRate,
        laborCost: qty * sym.laborHours * sym.laborRate,
        totalCost: qty * (sym.unitCost + sym.laborHours * sym.laborRate),
        priority: 'required', description: infraRules.dsp.description, mounting: sym.mounting
      });
    }
  }

  // Switches
  if (totalPoeDevices > 0) {
    let switchSymbol, qty;
    if (totalPoeDevices <= 20 && infraRules.switches.fallback) {
      switchSymbol = symbolMap[infraRules.switches.fallback.symbolId];
      qty = 1;
    } else {
      switchSymbol = symbolMap[infraRules.switches.symbolId];
      qty = evalFormula(infraRules.switches.formula.replace(/totalPoeDevices/g, totalPoeDevices), 0);
    }
    if (switchSymbol && qty > 0) {
      infraItems.push({
        symbolId: switchSymbol.id, name: switchSymbol.name, partNumber: switchSymbol.partNumber,
        manufacturer: switchSymbol.manufacturer, model: switchSymbol.model, category: switchSymbol.category,
        qty, unitCost: switchSymbol.unitCost, equipmentCost: qty * switchSymbol.unitCost,
        laborHours: switchSymbol.laborHours * qty, laborRate: switchSymbol.laborRate,
        laborCost: qty * switchSymbol.laborHours * switchSymbol.laborRate,
        totalCost: qty * (switchSymbol.unitCost + switchSymbol.laborHours * switchSymbol.laborRate),
        priority: 'required', description: infraRules.switches.description, mounting: switchSymbol.mounting
      });
    }
  }

  // Firewall
  const fwSym = symbolMap[infraRules.firewall.symbolId];
  if (fwSym) {
    infraItems.push({
      symbolId: fwSym.id, name: fwSym.name, partNumber: fwSym.partNumber,
      manufacturer: fwSym.manufacturer, model: fwSym.model, category: fwSym.category,
      qty: 1, unitCost: fwSym.unitCost, equipmentCost: fwSym.unitCost,
      laborHours: fwSym.laborHours, laborRate: fwSym.laborRate,
      laborCost: fwSym.laborHours * fwSym.laborRate,
      totalCost: fwSym.unitCost + fwSym.laborHours * fwSym.laborRate,
      priority: 'required', description: infraRules.firewall.description, mounting: fwSym.mounting
    });
  }

  // NVR
  if (totalCameras > 0) {
    const qty = evalFormula(infraRules.nvr.formula.replace(/totalCameras/g, totalCameras), 0);
    const sym = symbolMap[infraRules.nvr.symbolId];
    if (sym && qty > 0) {
      infraItems.push({
        symbolId: sym.id, name: sym.name, partNumber: sym.partNumber,
        manufacturer: sym.manufacturer, model: sym.model, category: sym.category,
        qty, unitCost: sym.unitCost, equipmentCost: qty * sym.unitCost,
        laborHours: sym.laborHours * qty, laborRate: sym.laborRate,
        laborCost: qty * sym.laborHours * sym.laborRate,
        totalCost: qty * (sym.unitCost + sym.laborHours * sym.laborRate),
        priority: 'required', description: infraRules.nvr.description, mounting: sym.mounting
      });
    }
  }

  // Access Panels
  if (totalReaders > 0) {
    const qty = evalFormula(infraRules.accessPanels.formula.replace(/totalReaders/g, totalReaders), 0);
    const sym = symbolMap[infraRules.accessPanels.symbolId];
    if (sym && qty > 0) {
      infraItems.push({
        symbolId: sym.id, name: sym.name, partNumber: sym.partNumber,
        manufacturer: sym.manufacturer, model: sym.model, category: sym.category,
        qty, unitCost: sym.unitCost, equipmentCost: qty * sym.unitCost,
        laborHours: sym.laborHours * qty, laborRate: sym.laborRate,
        laborCost: qty * sym.laborHours * sym.laborRate,
        totalCost: qty * (sym.unitCost + sym.laborHours * sym.laborRate),
        priority: 'required', description: infraRules.accessPanels.description, mounting: sym.mounting
      });
    }
  }

  return {
    zone: 'Infrastructure (Auto-Calculated)',
    type: 'infrastructure',
    label: 'Backbone Equipment',
    sqft: null,
    items: infraItems,
    counts: { totalSpeakers, totalZones, totalPoeDevices, totalCameras, totalReaders, totalDisplays },
    totals: infraItems.reduce((acc, item) => ({
      equipment: acc.equipment + item.equipmentCost,
      labor: acc.labor + item.laborCost,
      total: acc.total + item.totalCost,
      laborHours: acc.laborHours + item.laborHours
    }), { equipment: 0, labor: 0, total: 0, laborHours: 0 })
  };
}

/**
 * Generate complete BOQ from zone list
 */
function generateBOQ(projectName, zones) {
  const zoneResults = zones.map(processZone);
  const infrastructure = calculateInfrastructure(zoneResults);
  
  const allSections = [...zoneResults, infrastructure];
  
  const grandTotal = allSections.reduce((acc, section) => ({
    equipment: acc.equipment + section.totals.equipment,
    labor: acc.labor + section.totals.labor,
    total: acc.total + section.totals.total,
    laborHours: acc.laborHours + section.totals.laborHours
  }), { equipment: 0, labor: 0, total: 0, laborHours: 0 });

  const markup = rulesEngine.globalDefaults.markupPercent;
  const equipmentWithMarkup = grandTotal.equipment * (1 + markup / 100);

  return {
    project: projectName,
    generatedAt: new Date().toISOString(),
    zones: zoneResults,
    infrastructure,
    summary: {
      zoneCount: zones.length,
      totalSqft: zones.reduce((sum, z) => sum + z.sqft, 0),
      equipmentSubtotal: grandTotal.equipment,
      markupPercent: markup,
      equipmentWithMarkup,
      laborSubtotal: grandTotal.labor,
      totalLaborHours: grandTotal.laborHours,
      projectTotal: equipmentWithMarkup + grandTotal.labor
    }
  };
}

// --- DEMO RUN ---
const demoProject = generateBOQ('Oceanview Resort & Spa', [
  { name: 'Main Lobby', type: 'lobby', sqft: 2400 },
  { name: 'Fitness Center', type: 'fitness', sqft: 1200 },
  { name: 'Ballroom A', type: 'ballroom', sqft: 4500 },
  { name: 'Restaurant - Ocean Grill', type: 'restaurant', sqft: 1800 },
  { name: 'Lobby Bar', type: 'bar', sqft: 800 },
  { name: 'Pool Deck', type: 'pool', sqft: 3000 },
  { name: 'Spa & Wellness', type: 'spa', sqft: 1500 },
  { name: 'Meeting Room 1', type: 'meeting', sqft: 450 },
  { name: 'Meeting Room 2', type: 'meeting', sqft: 350 },
  { name: 'Main Corridor L1', type: 'corridor', sqft: 2000 },
  { name: 'BOH / Kitchen Corridor', type: 'back-of-house', sqft: 1500 },
  { name: 'Parking Garage', type: 'parking', sqft: 15000 }
]);

// Output results
console.log('\n' + '='.repeat(70));
console.log(`  PROJECT: ${demoProject.project}`);
console.log(`  Generated: ${demoProject.generatedAt}`);
console.log('='.repeat(70));

demoProject.zones.forEach(zone => {
  console.log(`\n--- ${zone.zone} (${zone.label}) — ${zone.sqft.toLocaleString()} sq ft ---`);
  zone.items.forEach(item => {
    console.log(`  ${item.qty}x ${item.name} (${item.partNumber}) — $${item.equipmentCost.toLocaleString()} equip + $${item.laborCost.toLocaleString()} labor`);
  });
  console.log(`  Zone Total: $${zone.totals.total.toLocaleString()}`);
});

console.log(`\n--- ${demoProject.infrastructure.zone} ---`);
console.log(`  Device counts: ${JSON.stringify(demoProject.infrastructure.counts)}`);
demoProject.infrastructure.items.forEach(item => {
  console.log(`  ${item.qty}x ${item.name} (${item.partNumber}) — $${item.equipmentCost.toLocaleString()} equip + $${item.laborCost.toLocaleString()} labor`);
});

console.log('\n' + '='.repeat(70));
console.log('  SUMMARY');
console.log('='.repeat(70));
const s = demoProject.summary;
console.log(`  Zones: ${s.zoneCount} | Total Area: ${s.totalSqft.toLocaleString()} sq ft`);
console.log(`  Equipment Subtotal:    $${s.equipmentSubtotal.toLocaleString()}`);
console.log(`  Markup (${s.markupPercent}%):          $${(s.equipmentWithMarkup - s.equipmentSubtotal).toLocaleString()}`);
console.log(`  Equipment w/ Markup:   $${s.equipmentWithMarkup.toLocaleString()}`);
console.log(`  Labor Subtotal:        $${s.laborSubtotal.toLocaleString()}`);
console.log(`  Total Labor Hours:     ${s.totalLaborHours.toLocaleString()} hrs`);
console.log(`  ─────────────────────────────────`);
console.log(`  PROJECT TOTAL:         $${s.projectTotal.toLocaleString()}`);
console.log('='.repeat(70));

// Save full output
fs.writeFileSync(
  path.join(__dirname, '../data/demo-boq-output.json'),
  JSON.stringify(demoProject, null, 2)
);
console.log('\nFull BOQ saved to data/demo-boq-output.json');

module.exports = { generateBOQ, processZone, calculateInfrastructure };
